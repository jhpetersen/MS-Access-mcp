using System.Data.OleDb;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace MS.Access.MCP.Interop
{
    public class AccessInteropService : IDisposable
    {
        private OleDbConnection? _oleDbConnection;
        private string? _currentDatabasePath;
        private bool _disposed = false;
        // Single shared Access COM instance – opened in Connect(), closed in Disconnect()
        private dynamic? _accessApp;
        // Original startup values saved at Connect, restored at Disconnect
        private string _savedStartupForm = "";
        private string _savedStartupMacro = "";

        [DllImport("user32.dll")] static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr hWnd);
        [DllImport("user32.dll")] static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
        [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr hWnd, System.Text.StringBuilder lpClassName, int nMaxCount);
        [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
        [DllImport("user32.dll")] static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowTextLength(IntPtr hWnd);
        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
        private const int SW_HIDE = 0;
        private const uint WM_COMMAND = 0x0111;
        private const uint BM_CLICK = 0x00F5;
        private const int IDOK = 1;

        #region 1. Connection Management

        public void Connect(string databasePath)
        {
            if (!File.Exists(databasePath))
                throw new FileNotFoundException($"Database file not found: {databasePath}");

            _currentDatabasePath = databasePath;

            // Step 1: Use DAO in exclusive mode to read and clear startup properties
            //         BEFORE any other connection opens the file.
            //         DAO Properties must be iterated via index (not foreach) for COM interop.
            ReadAndClearStartupPropertiesViaDao(databasePath);

            // Step 2: OleDb connection (data queries) — optional; COM works without it
            try
            {
                var connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};User Id=Admin;Password=;Jet OLEDB:System database=;";
                _oleDbConnection = new OleDbConnection(connectionString);
                _oleDbConnection.Open();
            }
            catch
            {
                // OleDb may be unavailable in some host environments.
                // VBA / form operations only need the COM instance below.
                _oleDbConnection = null;
            }

            // Step 3: Access COM instance.
            // During OpenCurrentDatabase we use AutomationSecurity=3 (ForceDisable) to suppress
            // all VBA compile/macro dialogs.  Immediately after the database is open we switch to
            // AutomationSecurity=1 (Low) so that compile_vba and other operations work normally.
            var appType = Type.GetTypeFromProgID("Access.Application")
                ?? throw new InvalidOperationException("Microsoft Access ist nicht installiert.");

            _accessApp = Activator.CreateInstance(appType)!;

            HideAccessWindow();
            _accessApp.Visible = false;

            // Use ForceDisable (3) during OpenCurrentDatabase so that Access suppresses all
            // VBA macro/compile activity silently – no dialogs can block the call.
            // We switch to Low (1 = msoAutomationSecurityLow) immediately after the database
            // is open so that compile_vba and other VBA operations work normally.
            _accessApp.AutomationSecurity = 3; // msoAutomationSecurityForceDisable – suppress dialogs during open

            // Start a watcher with pid=0 (no PID filter) to handle any non-VBA startup dialogs
            // (e.g. "falscher Verweis" / broken-reference warnings).  Using pid=0 means the
            // watcher accepts ALL visible #32770 dialogs regardless of which sub-process owns
            // them, which is safe here because we own this Access instance exclusively.
            using var startupCts = new System.Threading.CancellationTokenSource();
            var startupSink = new System.Collections.Generic.List<string>();
            var startupWatcher = System.Threading.Tasks.Task.Run(() =>
                WatchAccessDialogs(0, startupSink, startupCts.Token));

            _accessApp.OpenCurrentDatabase(databasePath, false, "");

            // Allow a moment for any post-open dialogs (broken-reference warnings etc.) to appear
            // and be dismissed by the watcher.  Keep AutomationSecurity=3 here so that no VBA
            // close-event code fires while we close startup forms — otherwise the compile error
            // dialog would re-appear when DoCmd.Close triggers Form_Close events.
            System.Threading.Thread.Sleep(1500);
            startupCts.Cancel();
            startupWatcher.Wait(1000);

            // Close any startup forms while VBA is still disabled – clean, no event dialogs.
            HideAccessWindow();
            // Belt-and-suspenders: also clear via SetOption (Access-native, persists on CloseCurrentDatabase)
            try { _accessApp.SetOption("Startup Form", ""); } catch { }
            try { _accessApp.SetOption("Startup Macro", ""); } catch { }
            CloseAllOpenForms(saveChanges: false);

            // Only now re-enable VBA so that compile_vba and subsequent operations work normally.
            _accessApp.AutomationSecurity = 1; // msoAutomationSecurityLow
        }

        private void ReadAndClearStartupPropertiesViaDao(string databasePath)
        {
            var daoType = Type.GetTypeFromProgID("DAO.DBEngine.120")
                ?? Type.GetTypeFromProgID("DAO.DBEngine.36")
                ?? Type.GetTypeFromProgID("DAO.DBEngine");
            if (daoType == null) return;

            dynamic daoEngine = Activator.CreateInstance(daoType)!;
            // Exclusive=true so we can write properties reliably
            dynamic daoDb = daoEngine.OpenDatabase(databasePath, true, false);
            try
            {
                _savedStartupForm = "";
                _savedStartupMacro = "";

                // DAO Properties collections must be iterated by INDEX in C# COM interop.
                // foreach over COM IEnumVARIANT does not reliably enumerate DAO collections.
                int count = (int)daoDb.Properties.Count;
                for (int i = 0; i < count; i++)
                {
                    try
                    {
                        dynamic prop = daoDb.Properties[i];
                        string pname = (string)prop.Name;
                        if (pname == "StartupForm")
                        {
                            _savedStartupForm = Nz(prop.Value);
                            prop.Value = "";
                        }
                        else if (pname == "StartupMacro")
                        {
                            _savedStartupMacro = Nz(prop.Value);
                            prop.Value = "";
                        }
                    }
                    catch { }
                }
            }
            finally
            {
                daoDb.Close();
                Marshal.ReleaseComObject(daoDb);
                Marshal.ReleaseComObject(daoEngine);
                // Ensure all COM wrappers are released before Access opens the file
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
        }

        private static string Nz(object? val) =>
            val == null || val is DBNull ? "" : val.ToString() ?? "";

        private void HideAccessWindow()
        {
            try
            {
                if (_accessApp == null) return;
                long hwndLong = (long)_accessApp.hWndAccessApp;
                if (hwndLong != 0)
                    ShowWindow(new IntPtr(hwndLong), SW_HIDE);
            }
            catch { }
        }

        private void RestoreStartupPropertiesViaDao()
        {
            if (_currentDatabasePath == null) return;
            if (_savedStartupForm == "" && _savedStartupMacro == "") return;

            var daoType = Type.GetTypeFromProgID("DAO.DBEngine.120")
                ?? Type.GetTypeFromProgID("DAO.DBEngine.36")
                ?? Type.GetTypeFromProgID("DAO.DBEngine");
            if (daoType == null) return;

            dynamic daoEngine = Activator.CreateInstance(daoType)!;
            dynamic daoDb = daoEngine.OpenDatabase(_currentDatabasePath, true, false);
            try
            {
                int count = (int)daoDb.Properties.Count;
                for (int i = 0; i < count; i++)
                {
                    try
                    {
                        dynamic prop = daoDb.Properties[i];
                        string pname = (string)prop.Name;
                        if (pname == "StartupForm" && _savedStartupForm != "")
                            prop.Value = _savedStartupForm;
                        else if (pname == "StartupMacro" && _savedStartupMacro != "")
                            prop.Value = _savedStartupMacro;
                    }
                    catch { }
                }
            }
            finally
            {
                daoDb.Close();
                Marshal.ReleaseComObject(daoDb);
                Marshal.ReleaseComObject(daoEngine);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        public void Disconnect()
        {
            // Restore startup options in-memory before closing (Access writes them to disk on CloseCurrentDatabase)
            try
            {
                if (_accessApp != null && _savedStartupForm != "")
                    _accessApp.SetOption("Startup Form", _savedStartupForm);
                if (_accessApp != null && _savedStartupMacro != "")
                    _accessApp.SetOption("Startup Macro", _savedStartupMacro);
            }
            catch { }

            // Quit with acQuitSaveAll (1) — saves all pending changes (VBA, forms, data)
            // without showing any save dialogs. Avoids CloseCurrentDatabase() which can
            // hang if Access is silently prompting about unsaved design objects.
            try { _accessApp?.Quit(1); } catch { }
            if (_accessApp != null) { Marshal.ReleaseComObject(_accessApp); _accessApp = null; }

            GC.Collect();
            GC.WaitForPendingFinalizers();

            // Extra safety: also restore via DAO in case SetOption didn't persist
            RestoreStartupPropertiesViaDao();

            _oleDbConnection?.Close();
            _oleDbConnection?.Dispose();
            _oleDbConnection = null;
            _currentDatabasePath = null;
        }

        private void EnsureAccessApp()
        {
            if (_accessApp == null)
                throw new InvalidOperationException("Not connected to database");
        }

        private void CloseAllOpenForms(bool saveChanges = false)
        {
            // acForm = 2, acSaveYes = 1, acSaveNo = 2
            const int acForm = 2;
            int saveMode = saveChanges ? 1 : 2;
            try
            {
                dynamic forms = _accessApp!.Forms;
                int count = (int)forms.Count;
                for (int i = count - 1; i >= 0; i--)
                {
                    try
                    {
                        string name = (string)forms[i].Name;
                        _accessApp.DoCmd.Close(acForm, name, saveMode);
                    }
                    catch { }
                }
            }
            catch { }
        }

        // Resolve VBComponent by module name.
        // Priority: Form_<name> > exact match > Report_<name>
        // This ensures form class modules take priority over any accidental standard module
        // with the same name (standard modules cannot use 'Me', causing compile errors).
        // Opens the form in design view so that its class module becomes visible in VBE.
        // Returns true if we actually opened it (caller must close with CloseFormDesignIfOpened).
        private bool EnsureFormModuleVisible(dynamic vbProject, string formName)
        {
            // Already accessible?
            try { var _ = vbProject.VBComponents("Form_" + formName); return false; } catch { }
            // Try to open in design view so Access creates / exposes the class module
            try
            {
                _accessApp!.DoCmd.OpenForm(formName, 1); // 1 = acDesign
                System.Threading.Thread.Sleep(400);
                return true;
            }
            catch { return false; }
        }

        private void CloseFormDesignIfOpened(string formName)
        {
            try { _accessApp!.DoCmd.Close(2, formName, 1); } // acForm=2, acSaveYes=1
            catch { }
        }

        // Fallback deletion when ProcStartLine cannot parse the proc name
        // (e.g. name contains non-identifier chars from prior encoding corruption).
        private void TryDeleteProcByLineScan(dynamic codeModule, string procedureName)
        {
            int total;
            try { total = (int)codeModule.CountOfLines; } catch { return; }
            string needle = procedureName.ToLowerInvariant();
            int startLine = -1;
            for (int i = 1; i <= total; i++)
            {
                string line;
                try { line = ((string)codeModule.Lines(i, 1)).Trim().ToLowerInvariant(); } catch { continue; }
                if ((line.Contains("sub ") || line.Contains("function ") || line.Contains("property ")) &&
                    line.Contains(needle))
                {
                    startLine = i;
                    break;
                }
            }
            if (startLine == -1) return;
            int endLine = -1;
            for (int i = startLine + 1; i <= total; i++)
            {
                string line;
                try { line = ((string)codeModule.Lines(i, 1)).Trim().ToLowerInvariant(); } catch { continue; }
                if (line == "end sub" || line == "end function" || line.StartsWith("end property"))
                {
                    endLine = i;
                    break;
                }
            }
            if (endLine == -1) return;
            try { codeModule.DeleteLines(startLine, endLine - startLine + 1); } catch { }
        }

        private dynamic GetVBComponent(string moduleName)
        {
            EnsureAccessApp();
            dynamic vbProject = _accessApp!.VBE.VBProjects(1);
            dynamic? component = null;

            // If already fully prefixed, resolve directly
            if (moduleName.StartsWith("Form_", StringComparison.OrdinalIgnoreCase) ||
                moduleName.StartsWith("Report_", StringComparison.OrdinalIgnoreCase))
            {
                try { component = vbProject.VBComponents(moduleName); } catch { }
                if (component == null) throw new ArgumentException($"Modul '{moduleName}' nicht gefunden.");
                return component;
            }

            // Try Form_ prefix first (class module beats same-named standard module)
            try { component = vbProject.VBComponents("Form_" + moduleName); } catch { }
            if (component == null)
            {
                // HasModule may be False — open in design view to expose the class module
                bool opened = EnsureFormModuleVisible(vbProject, moduleName);
                try { component = vbProject.VBComponents("Form_" + moduleName); } catch { }
                if (opened) CloseFormDesignIfOpened(moduleName);
            }
            if (component == null)
                try { component = vbProject.VBComponents(moduleName); } catch { }
            if (component == null)
                try { component = vbProject.VBComponents("Report_" + moduleName); } catch { }
            if (component == null)
                throw new ArgumentException($"Modul '{moduleName}' nicht gefunden.");
            return component;
        }

        // Delete a VBA module by name (removes spurious standard modules, etc.)
        public void DeleteVBAProcedure(string projectName, string moduleName, string procedureName)
        {
            EnsureAccessApp();
            dynamic vbProject = _accessApp!.VBE.VBProjects(1);
            dynamic component = GetVBComponent(moduleName);
            dynamic codeModule = component.CodeModule;

            bool deleted = false;
            try
            {
                int procStart = (int)codeModule.ProcStartLine(procedureName, 0);
                int procCount = (int)codeModule.ProcCountLines(procedureName, 0);
                if (procStart > 0 && procCount > 0)
                {
                    codeModule.DeleteLines(procStart, procCount);
                    deleted = true;
                }
            }
            catch { }
            if (!deleted)
            {
                TryDeleteProcByLineScan(codeModule, procedureName);
                // If still not found, just return silently (idempotent)
            }
        }

        public void DeleteVBAModule(string projectName, string moduleName)
        {
            EnsureAccessApp();
            dynamic vbProject = _accessApp!.VBE.VBProjects(1);
            dynamic? component = null;
            // For deletion, use EXACT match only — never delete a form class module by accident
            try { component = vbProject.VBComponents(moduleName); } catch { }
            if (component == null)
                throw new ArgumentException($"Standard-Modul '{moduleName}' nicht gefunden.");
            // Safety: only allow deleting standard modules (type 1), not form/report class modules
            int compType = 0;
            try { compType = (int)component.Type; } catch { }
            if (compType != 1)
                throw new InvalidOperationException($"'{moduleName}' ist kein Standard-Modul (Typ {compType}). Nur Standard-Module können so gelöscht werden.");
            vbProject.VBComponents.Remove(component);
        }

        public bool IsConnected => _accessApp != null ||
            (_oleDbConnection?.State == System.Data.ConnectionState.Open);

        #endregion

        #region 2. Data Access Object Models

        public List<TableInfo> GetTables()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var tables = new List<TableInfo>();

            // Prefer OleDb schema if available; fall back to DAO when OleDb is unavailable
            // (e.g. when Access already has the file open and blocks an OleDb connection).
            if (_oleDbConnection != null)
            {
                var schema = _oleDbConnection.GetSchema("Tables");
                foreach (System.Data.DataRow row in schema.Rows)
                {
                    var tableType = row["TABLE_TYPE"]?.ToString() ?? "";
                    var tableName = row["TABLE_NAME"]?.ToString() ?? "";
                    if (!string.IsNullOrEmpty(tableName) && !tableName.StartsWith("~")
                        && tableType == "TABLE")
                    {
                        tables.Add(new TableInfo
                        {
                            Name = tableName,
                            Fields = new List<FieldInfo>(),
                            RecordCount = 0
                        });
                    }
                }
            }
            else
            {
                // DAO fallback — works even when Access holds an exclusive lock
                var daoType = Type.GetTypeFromProgID("DAO.DBEngine.120")
                    ?? Type.GetTypeFromProgID("DAO.DBEngine.36")
                    ?? Type.GetTypeFromProgID("DAO.DBEngine")
                    ?? throw new InvalidOperationException("DAO.DBEngine COM-Klasse nicht gefunden.");
                dynamic engine = Activator.CreateInstance(daoType)!;
                dynamic db = engine.OpenDatabase(_currentDatabasePath, false, true);
                try
                {
                    foreach (dynamic td in db.TableDefs)
                    {
                        var tableName = (string)td.Name;
                        // Skip system tables (MSys*) and temp objects (~*)
                        if (tableName.StartsWith("MSys") || tableName.StartsWith("~"))
                            continue;
                        tables.Add(new TableInfo
                        {
                            Name = tableName,
                            Fields = new List<FieldInfo>(),
                            RecordCount = 0
                        });
                    }
                }
                finally { db.Close(); }
            }

            return tables;
        }

        public List<QueryInfo> GetQueries()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var queries = new List<QueryInfo>();

            // Use DAO to enumerate all QueryDefs (works even without OleDb connection)
            var daoType = Type.GetTypeFromProgID("DAO.DBEngine.120")
                ?? Type.GetTypeFromProgID("DAO.DBEngine.36")
                ?? Type.GetTypeFromProgID("DAO.DBEngine")
                ?? throw new InvalidOperationException("DAO nicht gefunden");
            dynamic engine = Activator.CreateInstance(daoType)!;
            dynamic db = engine.OpenDatabase(_currentDatabasePath, false, true);
            try
            {
                foreach (dynamic qd in db.QueryDefs)
                {
                    string name = (string)qd.Name;
                    // Skip hidden system queries (names starting with ~)
                    if (name.StartsWith("~")) continue;
                    queries.Add(new QueryInfo
                    {
                        Name = name,
                        SQL = "",
                        Type = "Query"
                    });
                }
            }
            finally
            {
                db.Close();
                Marshal.ReleaseComObject(db);
            }

            return queries;
        }

        public List<RelationshipInfo> GetRelationships()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var relationships = new List<RelationshipInfo>();
            var daoType = Type.GetTypeFromProgID("DAO.DBEngine.120")
                ?? Type.GetTypeFromProgID("DAO.DBEngine.36")
                ?? throw new InvalidOperationException("DAO nicht gefunden");
            dynamic engine = Activator.CreateInstance(daoType)!;
            dynamic db = engine.OpenDatabase(_currentDatabasePath, false, true);
            try
            {
                foreach (dynamic rel in db.Relations)
                {
                    relationships.Add(new RelationshipInfo
                    {
                        Name = (string)rel.Name,
                        Table = (string)rel.Table,
                        ForeignTable = (string)rel.ForeignTable,
                        Attributes = rel.Attributes.ToString()
                    });
                }
            }
            finally { db.Close(); }
            return relationships;
        }

        private static string MapFieldTypeToJet(string typeName, int size)
        {
            return typeName.ToLowerInvariant() switch
            {
                "autonumber" or "counter" => "COUNTER",
                "long" or "integer" or "int" => "LONG",
                "short" or "smallint" => "SHORT",
                "text" or "char" or "varchar" => size > 0 ? $"TEXT({size})" : "TEXT(255)",
                "memo" or "longtext" => "MEMO",
                "datetime" or "date" => "DATETIME",
                "boolean" or "bit" or "yesno" => "BIT",
                "double" or "float" => "DOUBLE",
                "currency" or "money" => "CURRENCY",
                "single" or "real" => "SINGLE",
                "byte" => "BYTE",
                _ => typeName   // pass through unknown types verbatim
            };
        }

        public void CreateTable(string tableName, List<FieldInfo> fields)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            EnsureAccessApp();

            var fieldDefinitions = new List<string>();
            foreach (var field in fields)
            {
                var jetType = MapFieldTypeToJet(field.Type, field.Size);
                var fieldDef = $"[{field.Name}] {jetType}";
                if (field.Required)
                    fieldDef += " NOT NULL";
                fieldDefinitions.Add(fieldDef);
            }

            var createSql = $"CREATE TABLE [{tableName}] ({string.Join(", ", fieldDefinitions)})";
            _accessApp!.CurrentDb().Execute(createSql);
        }

        public void DeleteTable(string tableName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            EnsureAccessApp();
            _accessApp!.CurrentDb().Execute($"DROP TABLE [{tableName}]");
        }

        #endregion

        #region 3. COM Automation (Simplified)

        public void LaunchAccess()
        {
            // This would require full COM interop - simplified for now
            Console.WriteLine("Access launch functionality requires full COM interop");
        }

        public void CloseAccess()
        {
            // Full cleanup: close OleDb connection, restore startup props, quit Access.
            // Reuse Disconnect() logic so nothing is left dangling.
            Disconnect();
        }

        public List<FormInfo> GetForms()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            var forms = new List<FormInfo>();
            foreach (var name in GetAccessObjectNamesViaDao("Forms"))
                forms.Add(new FormInfo { Name = name, FullName = name, Type = "Form" });
            return forms;
        }

        public List<ReportInfo> GetReports()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            var reports = new List<ReportInfo>();
            foreach (var name in GetAccessObjectNamesViaDao("Reports"))
                reports.Add(new ReportInfo { Name = name, FullName = name, Type = "Report" });
            return reports;
        }

        public List<MacroInfo> GetMacros()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var macros = new List<MacroInfo>();
            foreach (var name in GetAccessObjectNamesViaDao("Scripts"))
                macros.Add(new MacroInfo { Name = name, FullName = name, Type = "Macro" });
            return macros;
        }

        public List<ModuleInfo> GetModules()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var modules = new List<ModuleInfo>();
            var command2 = new OleDbCommand("SELECT Name FROM MSysObjects WHERE Type = -32761 ORDER BY Name", _oleDbConnection);
            using var reader2 = command2.ExecuteReader();
            while (reader2.Read())
            {
                modules.Add(new ModuleInfo
                {
                    Name = reader2["Name"]?.ToString() ?? "",
                    FullName = reader2["Name"]?.ToString() ?? "",
                    Type = "Module"
                });
            }
            return modules;
        }

        private IEnumerable<string> GetAccessObjectNamesViaDao(string containerName)
        {
            var daoType = Type.GetTypeFromProgID("DAO.DBEngine.120")
                ?? Type.GetTypeFromProgID("DAO.DBEngine.36")
                ?? Type.GetTypeFromProgID("DAO.DBEngine")
                ?? throw new InvalidOperationException("DAO.DBEngine COM-Klasse nicht gefunden. Ist Microsoft Access installiert?");

            dynamic engine = Activator.CreateInstance(daoType)!;
            dynamic db = engine.OpenDatabase(_currentDatabasePath, false, true);
            var names = new List<string>();
            try
            {
                dynamic container = db.Containers[containerName];
                foreach (dynamic doc in container.Documents)
                    names.Add((string)doc.Name);
            }
            finally { db.Close(); }
            return names;
        }

        public void OpenForm(string formName)
        {
            EnsureAccessApp();
            _accessApp!.DoCmd.OpenForm(formName, 0); // acNormal = 0
        }

        public void CloseForm(string formName)
        {
            EnsureAccessApp();
            const int acForm = 2;
            const int acSaveYes = 1;
            try { _accessApp!.DoCmd.Close(acForm, formName, acSaveYes); } catch { }
        }

        #endregion

        #region 4. VBA Extensibility (Simplified)

        public List<VBAProjectInfo> GetVBAProjects()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var projects = new List<VBAProjectInfo>();
            var moduleNames = GetAccessObjectNamesViaDao("Modules");
            var modules = moduleNames.Select(n => new VBAModuleInfo { Name = n, Type = "Module", HasCode = true }).ToList();
            projects.Add(new VBAProjectInfo
            {
                Name = "CurrentProject",
                Description = "Current Access Project",
                Modules = modules
            });
            return projects;
        }

        public string GetVBACode(string projectName, string moduleName)
        {
            dynamic component = GetVBComponent(moduleName);
            dynamic codeModule = component.CodeModule;
            int lineCount = (int)codeModule.CountOfLines;
            if (lineCount == 0) return "";
            return (string)codeModule.Lines(1, lineCount);
        }

        public void SetVBACode(string projectName, string moduleName, string code)
        {
            EnsureAccessApp();

            // --- Strategy A: Try to find the VBComponent via VBE directly -----------
            // Only works for standard modules and form modules that already have HasModule=True.
            dynamic vbProject = _accessApp!.VBE.VBProjects(1);
            dynamic? component = null;

            if (moduleName.StartsWith("Form_", StringComparison.OrdinalIgnoreCase) ||
                moduleName.StartsWith("Report_", StringComparison.OrdinalIgnoreCase))
            {
                try { component = vbProject.VBComponents(moduleName); } catch { }
            }
            else
            {
                try { component = vbProject.VBComponents("Form_" + moduleName); } catch { }
                if (component == null) try { component = vbProject.VBComponents(moduleName); } catch { }
                if (component == null) try { component = vbProject.VBComponents("Report_" + moduleName); } catch { }
            }

            if (component != null)
            {
                // Found via VBE — write code directly into the CodeModule
                dynamic codeModule = component.CodeModule;
                int lc = (int)codeModule.CountOfLines;
                if (lc > 0) codeModule.DeleteLines(1, lc);
                if (!string.IsNullOrEmpty(code)) codeModule.InsertLines(1, code);
                return;
            }

            // --- Strategy B: Form/Report with HasModule=False -------------------------
            // Open in Design view, access .Module property (this sets HasModule=True),
            // write code, then save+close.
            // Per MS docs: "When you use a method of the Module object or refer to the
            // Module property for a form in Design view, Access creates the associated
            // module and sets HasModule to True."
            bool isKnownForm = TrySetFormModuleCode(moduleName, code);
            if (isKnownForm) return;

            bool isKnownReport = TrySetReportModuleCode(moduleName, code);
            if (isKnownReport) return;

            // --- Strategy C: Not a form/report — create/replace standard module ------
            try { component = vbProject.VBComponents(moduleName); } catch { }
            if (component == null)
            {
                component = vbProject.VBComponents.Add(1); // vbext_ct_StdModule
                component.Name = moduleName;
            }
            dynamic cm = component.CodeModule;
            int lineCount = (int)cm.CountOfLines;
            if (lineCount > 0) cm.DeleteLines(1, lineCount);
            if (!string.IsNullOrEmpty(code)) cm.InsertLines(1, code);
        }

        /// <summary>
        /// Opens a form in design view, accesses .Module to force HasModule=True,
        /// writes the code, then saves and closes.  Returns false if the form is
        /// not found in the AllForms collection (i.e. it is not a form at all).
        /// </summary>
        private bool TrySetFormModuleCode(string formName, string code)
        {
            // Check if this name exists as a form by iterating AllForms
            bool exists = false;
            try
            {
                foreach (dynamic item in _accessApp!.CurrentProject.AllForms)
                {
                    if (string.Equals((string)item.Name, formName, StringComparison.OrdinalIgnoreCase))
                    { exists = true; break; }
                }
            }
            catch { }
            if (!exists) return false;

            const int acDesign = 1;
            const int acForm   = 2;
            const int acSaveYes = 1;

            _accessApp!.DoCmd.OpenForm(formName, acDesign);
            System.Threading.Thread.Sleep(600);

            try
            {
                // Accessing .Module on an open Design-view form forces HasModule=True
                dynamic frm = _accessApp.Forms(formName);
                frm.HasModule = true;
                dynamic mdl = frm.Module;
                int lc = (int)mdl.CountOfLines;
                if (lc > 0) mdl.DeleteLines(1, lc);
                if (!string.IsNullOrEmpty(code)) mdl.InsertLines(1, code);
            }
            finally
            {
                try { _accessApp!.DoCmd.Close(acForm, formName, acSaveYes); } catch { }
            }
            return true;
        }

        /// <summary>
        /// Same as TrySetFormModuleCode but for reports.
        /// </summary>
        private bool TrySetReportModuleCode(string reportName, string code)
        {
            bool exists = false;
            try
            {
                foreach (dynamic item in _accessApp!.CurrentProject.AllReports)
                {
                    if (string.Equals((string)item.Name, reportName, StringComparison.OrdinalIgnoreCase))
                    { exists = true; break; }
                }
            }
            catch { }
            if (!exists) return false;

            const int acDesign  = 1;
            const int acReport  = 3;
            const int acSaveYes = 1;

            _accessApp!.DoCmd.OpenReport(reportName, acDesign);
            System.Threading.Thread.Sleep(600);

            try
            {
                dynamic rpt = _accessApp.Reports(reportName);
                rpt.HasModule = true;
                dynamic mdl = rpt.Module;
                int lc = (int)mdl.CountOfLines;
                if (lc > 0) mdl.DeleteLines(1, lc);
                if (!string.IsNullOrEmpty(code)) mdl.InsertLines(1, code);
            }
            finally
            {
                try { _accessApp!.DoCmd.Close(acReport, reportName, acSaveYes); } catch { }
            }
            return true;
        }

        public void AddVBAProcedure(string projectName, string moduleName, string procedureName, string code)
        {
            EnsureAccessApp();

            dynamic vbProject = _accessApp!.VBE.VBProjects(1);
            dynamic? component = null;

            if (moduleName.StartsWith("Form_", StringComparison.OrdinalIgnoreCase) ||
                moduleName.StartsWith("Report_", StringComparison.OrdinalIgnoreCase))
            {
                try { component = vbProject.VBComponents(moduleName); } catch { }
            }
            else
            {
                try { component = vbProject.VBComponents("Form_" + moduleName); } catch { }
                if (component == null) try { component = vbProject.VBComponents(moduleName); } catch { }
                if (component == null) try { component = vbProject.VBComponents("Report_" + moduleName); } catch { }
            }

            if (component != null)
            {
                // Found via VBE — delete existing proc then append
                dynamic codeModule = component.CodeModule;
                bool procDeleted = false;
                try
                {
                    int procStart = (int)codeModule.ProcStartLine(procedureName, 0);
                    int procCount = (int)codeModule.ProcCountLines(procedureName, 0);
                    if (procStart > 0 && procCount > 0)
                    { codeModule.DeleteLines(procStart, procCount); procDeleted = true; }
                }
                catch { }
                if (!procDeleted) TryDeleteProcByLineScan(codeModule, procedureName);
                int lineCount = (int)codeModule.CountOfLines;
                string insertCode2 = lineCount > 0 ? "\r\n" + code : code;
                codeModule.InsertLines(lineCount + 1, insertCode2);
                return;
            }

            // Component not found via VBE — it may be a form/report with HasModule=False.
            // Open in Design view, force HasModule=True via .Module, write proc, save+close.
            if (TryAddProcToFormModule(moduleName, procedureName, code)) return;
            if (TryAddProcToReportModule(moduleName, procedureName, code)) return;

            // Not a form/report — create a standard module
            component = vbProject.VBComponents.Add(1); // vbext_ct_StdModule
            component.Name = moduleName;
            dynamic cm = component.CodeModule;
            int lc = (int)cm.CountOfLines;
            string ic = lc > 0 ? "\r\n" + code : code;
            cm.InsertLines(lc + 1, ic);
        }

        private bool TryAddProcToFormModule(string formName, string procedureName, string code)
        {
            bool exists = false;
            try { foreach (dynamic item in _accessApp!.CurrentProject.AllForms)
                { if (string.Equals((string)item.Name, formName, StringComparison.OrdinalIgnoreCase)) { exists = true; break; } } }
            catch { }
            if (!exists) return false;

            _accessApp!.DoCmd.OpenForm(formName, 1); // acDesign
            System.Threading.Thread.Sleep(600);
            try
            {
                dynamic frm = _accessApp.Forms(formName);
                frm.HasModule = true;
                dynamic mdl = frm.Module;
                TryDeleteProcByLineScan(mdl, procedureName);
                int lc = (int)mdl.CountOfLines;
                string ic = lc > 0 ? "\r\n" + code : code;
                mdl.InsertLines(lc + 1, ic);
            }
            finally { try { _accessApp!.DoCmd.Close(2, formName, 1); } catch { } }
            return true;
        }

        private bool TryAddProcToReportModule(string reportName, string procedureName, string code)
        {
            bool exists = false;
            try { foreach (dynamic item in _accessApp!.CurrentProject.AllReports)
                { if (string.Equals((string)item.Name, reportName, StringComparison.OrdinalIgnoreCase)) { exists = true; break; } } }
            catch { }
            if (!exists) return false;

            _accessApp!.DoCmd.OpenReport(reportName, 1); // acDesign
            System.Threading.Thread.Sleep(600);
            try
            {
                dynamic rpt = _accessApp.Reports(reportName);
                rpt.HasModule = true;
                dynamic mdl = rpt.Module;
                TryDeleteProcByLineScan(mdl, procedureName);
                int lc = (int)mdl.CountOfLines;
                string ic = lc > 0 ? "\r\n" + code : code;
                mdl.InsertLines(lc + 1, ic);
            }
            finally { try { _accessApp!.DoCmd.Close(3, reportName, 1); } catch { } }
            return true;
        }

        public CompileResult CompileVBA()
        {
            EnsureAccessApp();

            // Determine the PID of the Access process so the dialog watcher can
            // filter windows belonging to exactly this instance.
            uint accessPid = 0;
            try
            {
                long hwndLong = (long)_accessApp!.hWndAccessApp;
                if (hwndLong != 0)
                    GetWindowThreadProcessId(new IntPtr(hwndLong), out accessPid);
            }
            catch { }

            var capturedErrors = new System.Collections.Generic.List<string>();

            // A background thread watches for the VBE error dialog (class "#32770").
            // vbProject.Compile() blocks until the user dismisses the dialog, so the
            // watcher must dismiss it for us and capture the error text.
            // Use pid=0 fallback so that dialogs from sub-threads/VBE host are caught too.
            using var cts = new System.Threading.CancellationTokenSource();
            var watchTask = System.Threading.Tasks.Task.Run(() =>
                WatchAccessDialogs(accessPid != 0 ? accessPid : 0, capturedErrors, cts.Token));

            // VBProject.Compile() is not reliably reachable via C# dynamic late binding
            // (IDispatch may not expose it).  We use Type.InvokeMember to call it directly
            // through reflection, bypassing the C# dynamic binder entirely.
            // Fallback: VBE CommandBars control ID 578 = "Alle Module kompilieren".
            bool isCompiled = false;
            string? errorModule = null;
            int? errorLine = null;

            bool compileCalled = false;
            try
            {
                // Ensure VBE window stays hidden before and after compile.
                try { _accessApp!.VBE.MainWindow.Visible = false; } catch { }

                // Get VBProject as a raw COM object and invoke Compile() via reflection.
                object vbProject = _accessApp!.VBE.VBProjects.Item(1);
                vbProject.GetType().InvokeMember(
                    "Compile",
                    System.Reflection.BindingFlags.InvokeMethod |
                    System.Reflection.BindingFlags.Public |
                    System.Reflection.BindingFlags.Instance,
                    null, vbProject, null);
                compileCalled = true;
                isCompiled = true; // no exception → success
            }
            catch (System.Reflection.TargetInvocationException tie)
                when (tie.InnerException is System.Runtime.InteropServices.COMException)
            {
                // Compile() threw via InvokeMember → compile error dialog was shown.
                compileCalled = true;
                try { errorModule = (string)_accessApp!.VBE.SelectedVBComponent.Name; } catch { }
                try { errorLine = (int)_accessApp!.VBE.ActiveCodePane.TopLine; } catch { }
            }
            catch { /* InvokeMember failed entirely – try CommandBars fallback below */ }

            if (!compileCalled)
            {
                // Fallback: trigger compile via VBE CommandBars control (ID 578).
                // Keep VBE window hidden to avoid flicker — CommandBars work without it being visible.
                try
                {
                    dynamic ctrl = _accessApp!.VBE.CommandBars.FindControl(
                        Type: 1 /*msoControlButton*/, Id: 578);
                    ctrl.Execute();
                    compileCalled = true;
                    // CommandBars.Execute is async — wait up to 10 s for the dialog watcher
                    // to signal an error or for IsCompiled to flip true.
                    var cbDeadline = DateTime.UtcNow.AddSeconds(10);
                    while (DateTime.UtcNow < cbDeadline)
                    {
                        if (capturedErrors.Count > 0) break;
                        try { if ((bool)_accessApp!.IsCompiled) { isCompiled = true; break; } } catch { break; }
                        System.Threading.Thread.Sleep(100);
                    }
                    if (!isCompiled)
                    {
                        try { errorModule = (string)_accessApp!.VBE.SelectedVBComponent.Name; } catch { }
                        try { errorLine = (int)_accessApp!.VBE.ActiveCodePane.TopLine; } catch { }
                    }
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    try { errorModule = (string)_accessApp!.VBE.SelectedVBComponent.Name; } catch { }
                    try { errorLine = (int)_accessApp!.VBE.ActiveCodePane.TopLine; } catch { }
                }
                catch { capturedErrors.Add("Compile not accessible via VBE object model or CommandBars."); }
            }

            // Give watcher a brief moment to collect text from any dialog that just closed.
            System.Threading.Thread.Sleep(500);
            cts.Cancel();
            watchTask.Wait(2000);

            // Hide the VBE window if it became visible during compile.
            try { _accessApp!.VBE.MainWindow.Visible = false; } catch { }

            if (!isCompiled && capturedErrors.Count == 0)
            {
                capturedErrors.Add("Compile failed (error dialog text not captured; check ErrorModule/ErrorLine)");
            }

            // Deduplicate: the watcher may capture the same dialog text multiple times
            // if it iterates through EnumWindows several times before the dialog is dismissed.
            var uniqueErrors = capturedErrors
                .Distinct(System.StringComparer.Ordinal)
                .ToList();

            return new CompileResult
            {
                IsCompiled = isCompiled,
                Errors = uniqueErrors,
                ErrorModule = errorModule,
                ErrorLine = errorLine
            };
        }

        // Loops until cancellation, looking for modal Win32 dialogs belonging to
        // the given Access process. When found, collects the message text from all
        // Static child controls and closes the dialog via WM_COMMAND/IDOK.
        private void WatchAccessDialogs(
            uint accessPid,
            System.Collections.Generic.List<string> results,
            System.Threading.CancellationToken ct)
        {
            while (!ct.IsCancellationRequested)
            {
                EnumWindows((hWnd, _) =>
                {
                    if (ct.IsCancellationRequested) return false; // stop enumeration early

                    // Check that this top-level window belongs to our Access process.
                    GetWindowThreadProcessId(hWnd, out uint pid);
                    // If accessPid is 0 (PID detection failed), accept dialogs from any process.
                    if (accessPid != 0 && pid != accessPid) return true;

                    // Only handle visible dialogs (class "#32770" = standard MessageBox/Dialog).
                    var cls = new System.Text.StringBuilder(64);
                    GetClassName(hWnd, cls, cls.Capacity);
                    if (cls.ToString() != "#32770") return true;
                    if (!IsWindowVisible(hWnd)) return true;

                    // Collect text from all Static child controls (= the message label(s)).
                    var texts = new System.Collections.Generic.List<string>();
                    EnumChildWindows(hWnd, (child, _) =>
                    {
                        var childCls = new System.Text.StringBuilder(64);
                        GetClassName(child, childCls, childCls.Capacity);
                        if (childCls.ToString() == "Static")
                        {
                            var sb = new System.Text.StringBuilder(512);
                            if (GetWindowText(child, sb, sb.Capacity) > 0)
                            {
                                var txt = sb.ToString().Trim();
                                if (txt.Length > 0) texts.Add(txt);
                            }
                        }
                        return true;
                    }, IntPtr.Zero);

                    if (texts.Count > 0)
                    {
                        lock (results)
                            results.Add(string.Join(" | ", texts));
                    }

                    // Dismiss the dialog: find the OK/Beenden button and send BM_CLICK.
                    // SendMessage is synchronous — the dialog is guaranteed to be gone before
                    // this call returns, unlike PostMessage which is fire-and-forget.
                    IntPtr okButton = IntPtr.Zero;
                    EnumChildWindows(hWnd, (child, _) =>
                    {
                        var childCls2 = new System.Text.StringBuilder(32);
                        GetClassName(child, childCls2, childCls2.Capacity);
                        if (childCls2.ToString() == "Button")
                        {
                            var btnText = new System.Text.StringBuilder(64);
                            GetWindowText(child, btnText, btnText.Capacity);
                            string t = btnText.ToString().Trim();
                            // Match OK, &OK, Beenden, &Beenden (German Access)
                            if (t == "OK" || t == "&OK" || t.Contains("Beenden") || t.Contains("Ende"))
                            {
                                okButton = child;
                                return false; // stop enumeration
                            }
                        }
                        return true;
                    }, IntPtr.Zero);

                    if (okButton != IntPtr.Zero)
                        SendMessage(okButton, BM_CLICK, IntPtr.Zero, IntPtr.Zero);
                    else
                        SendMessage(hWnd, WM_COMMAND, new IntPtr(IDOK), IntPtr.Zero);

                    return true;
                }, IntPtr.Zero);

                if (!ct.IsCancellationRequested)
                    System.Threading.Thread.Sleep(50);
            }
        }

        #endregion

        #region 5. System Table Metadata Access

        public List<SystemTableInfo> GetSystemTables()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var systemTables = new List<SystemTableInfo>();
            var schema = _oleDbConnection!.GetSchema("Tables");
            
            foreach (System.Data.DataRow row in schema.Rows)
            {
                var tableName = row["TABLE_NAME"].ToString();
                if (!string.IsNullOrEmpty(tableName) && (tableName.StartsWith("~") || tableName.StartsWith("MSys")))
                {
                    systemTables.Add(new SystemTableInfo
                    {
                        Name = tableName,
                        DateCreated = DateTime.Now, // Not available through OleDb
                        LastUpdated = DateTime.Now, // Not available through OleDb
                        RecordCount = GetTableRecordCount(tableName)
                    });
                }
            }

            return systemTables;
        }

        public List<MetadataInfo> GetObjectMetadata()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var metadata = new List<MetadataInfo>();
            var daoType = Type.GetTypeFromProgID("DAO.DBEngine.120")
                ?? Type.GetTypeFromProgID("DAO.DBEngine.36")
                ?? throw new InvalidOperationException("DAO nicht gefunden");
            dynamic engine = Activator.CreateInstance(daoType)!;
            dynamic db = engine.OpenDatabase(_currentDatabasePath, false, true);
            try
            {
                foreach (dynamic td in db.TableDefs)
                {
                    var name = (string)td.Name;
                    if (name.StartsWith("MSys")) continue;
                    var fields = new System.Text.StringBuilder();
                    foreach (dynamic f in td.Fields)
                        fields.Append($"{f.Name}({f.Type}), ");
                    metadata.Add(new MetadataInfo
                    {
                        Name = name,
                        Type = "Table",
                        Flags = fields.ToString().TrimEnd(',', ' '),
                        DateCreated = "",
                        DateModified = ""
                    });
                }
            }
            finally { db.Close(); }
            return metadata;
        }

        #endregion

        #region 6. Form & Control Discovery & Editing APIs (Simplified)

        public bool FormExists(string formName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            
            try
            {
                var command = new OleDbCommand("SELECT COUNT(*) FROM MSysObjects WHERE Name = ? AND Type = -32768", _oleDbConnection);
                command.Parameters.AddWithValue("@Name", formName);
                var count = Convert.ToInt32(command.ExecuteScalar());
                return count > 0;
            }
            catch
            {
                return false;
            }
        }

        public List<ControlInfo> GetFormControls(string formName)
        {
            EnsureAccessApp();
            const int acForm = 2;
            const int acDesign = 1;
            const int acSaveNo = 2;

            _accessApp!.DoCmd.OpenForm(formName, acDesign);
            try
            {
                return ReadControlsFromCollection(_accessApp.Forms(formName).Controls);
            }
            finally
            {
                try { _accessApp.DoCmd.Close(acForm, formName, acSaveNo); } catch { }
            }
        }

        private List<ControlInfo> ReadControlsFromCollection(dynamic controlsCollection)
        {
            var controls = new List<ControlInfo>();
            int count = 0;
            try { count = (int)controlsCollection.Count; } catch { return controls; }

            for (int i = 0; i < count; i++)
            {
                try
                {
                    dynamic ctrl = controlsCollection[i];
                    int ctrlType = 0;
                    try { ctrlType = (int)ctrl.ControlType; } catch { }

                    var info = new ControlInfo
                    {
                        Name = ctrl.Name?.ToString() ?? $"Control{i}",
                        Type = ctrlType switch
                        {
                            100 => "Label",
                            101 => "Rectangle",
                            104 => "CommandButton",
                            105 => "OptionButton",
                            106 => "ListBox",
                            109 => "TextBox",
                            110 => "CheckBox",
                            111 => "ComboBox",
                            112 => "SubForm",
                            118 => "Page",
                            119 => "TabControl",
                            122 => "Image",
                            _ => $"Type{ctrlType}"
                        }
                    };
                    try { info.Left = (int)ctrl.Left; } catch { }
                    try { info.Top = (int)ctrl.Top; } catch { }
                    try { info.Width = (int)ctrl.Width; } catch { }
                    try { info.Height = (int)ctrl.Height; } catch { }
                    try { info.Caption = ctrl.Caption?.ToString(); } catch { }
                    try { info.ControlSource = ctrl.ControlSource?.ToString(); } catch { }
                    try { info.SourceObject = ctrl.SourceObject?.ToString(); } catch { }
                    try { info.LinkChildFields = ctrl.LinkChildFields?.ToString(); } catch { }
                    try { info.LinkMasterFields = ctrl.LinkMasterFields?.ToString(); } catch { }
                    try { info.Visible = (bool)ctrl.Visible; } catch { info.Visible = true; }
                    try { info.Enabled = (bool)ctrl.Enabled; } catch { info.Enabled = true; }
                    try { info.BackColor = (int)ctrl.BackColor; } catch { }
                    try { info.ForeColor = (int)ctrl.ForeColor; } catch { }
                    try { info.FontBold = (bool)ctrl.FontBold; } catch { }
                    try { info.FontSize = (int)ctrl.FontSize; } catch { }
                    try { info.SpecialEffect = (int)ctrl.SpecialEffect != 0; } catch { }
                    controls.Add(info);
                }
                catch { }
            }
            return controls;
        }

        public ControlProperties GetControlProperties(string formName, string controlName)
        {
            EnsureAccessApp();
            const int acForm = 2;
            const int acDesign = 1;
            const int acSaveNo = 2;

            _accessApp!.DoCmd.OpenForm(formName, acDesign);
            try
            {
                dynamic frm = _accessApp.Forms(formName);
                dynamic ctrl = frm.Controls(controlName);
                int ctrlType = 0;
                try { ctrlType = (int)ctrl.ControlType; } catch { }

                var props = new ControlProperties
                {
                    Name = ctrl.Name?.ToString() ?? controlName,
                    Type = ctrlType switch
                    {
                        100 => "Label", 101 => "Rectangle", 104 => "CommandButton",
                        105 => "OptionButton", 106 => "ListBox", 109 => "TextBox",
                        110 => "CheckBox", 111 => "ComboBox", 112 => "SubForm",
                        118 => "Page", 119 => "TabControl", 122 => "Image",
                        _ => $"Type{ctrlType}"
                    }
                };
                try { props.Left = (int)ctrl.Left; } catch { }
                try { props.Top = (int)ctrl.Top; } catch { }
                try { props.Width = (int)ctrl.Width; } catch { }
                try { props.Height = (int)ctrl.Height; } catch { }
                try { props.Visible = (bool)ctrl.Visible; } catch { props.Visible = true; }
                try { props.Enabled = (bool)ctrl.Enabled; } catch { props.Enabled = true; }
                try { props.BackColor = (int)ctrl.BackColor; } catch { }
                try { props.ForeColor = (int)ctrl.ForeColor; } catch { }
                try { props.FontName = ctrl.FontName?.ToString() ?? ""; } catch { }
                try { props.FontSize = (int)ctrl.FontSize; } catch { }
                try { props.FontBold = (bool)ctrl.FontBold; } catch { }
                try { props.FontItalic = (bool)ctrl.FontItalic; } catch { }
                return props;
            }
            finally
            {
                try { _accessApp.DoCmd.Close(acForm, formName, acSaveNo); } catch { }
            }
        }

        public void SetControlProperty(string formName, string controlName, string propertyName, object value)
        {
            EnsureAccessApp();

            const int acForm = 2;
            const int acDesign = 1;
            const int acSaveYes = 1;

            // Open the form in design view
            _accessApp!.DoCmd.OpenForm(formName, acDesign);
            try
            {
                dynamic frm = _accessApp.Forms(formName);

                // Form-level property (controlName is null/"" or matches form name)
                if (string.IsNullOrEmpty(controlName) || controlName == formName)
                {
                    try { frm.GetType().InvokeMember(propertyName,
                        System.Reflection.BindingFlags.SetProperty, null, frm,
                        new object[] { value }); } catch { }
                }
                else
                {
                    dynamic ctrl = frm.Controls(controlName);
                    try { ctrl.GetType().InvokeMember(propertyName,
                        System.Reflection.BindingFlags.SetProperty, null, ctrl,
                        new object[] { value }); } catch { }
                }
            }
            finally
            {
                _accessApp.DoCmd.Close(acForm, formName, acSaveYes);
            }
        }

        #endregion

        #region 7. Persistence & Versioning

        public string ExportFormToText(string formName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var formData = new
            {
                Name = formName,
                ExportedAt = DateTime.UtcNow,
                Controls = GetFormControls(formName),
                VBA = GetVBACode("CurrentProject", formName)
            };

            return JsonSerializer.Serialize(formData, new JsonSerializerOptions { WriteIndented = true });
        }

        public void ImportFormFromText(string formData)
        {
            EnsureAccessApp();

            var formInfo = JsonSerializer.Deserialize<FormExportData>(formData);
            if (formInfo == null) throw new ArgumentException("Invalid form data");
            if (string.IsNullOrWhiteSpace(formInfo.Name)) throw new ArgumentException("Form name is required");

            const int acForm = 2;         // AcObjectType.acForm
            const int acSaveYes = 1;      // AcCloseSave.acSaveYes
            const int acDetail = 0;       // AcSection.acDetail
            const int acTextBox = 109;
            const int acLabel = 100;
            const int acCommandButton = 104;
            const int acSubform = 112;
            const int acRectangle = 101;

            // Delete existing form if present
            try { _accessApp!.DoCmd.DeleteObject(acForm, formInfo.Name); } catch { }

            // CreateForm() creates a new blank form in design view
            dynamic frm = _accessApp!.CreateForm();
            string tempName = (string)frm.Name;

            // Set form-level properties while in design view
            try { if (formInfo.RecordSource != null) frm.RecordSource = formInfo.RecordSource; } catch { }
            try { if (formInfo.DefaultView.HasValue) frm.DefaultView = formInfo.DefaultView.Value; } catch { }
            try { if (formInfo.Popup.HasValue) frm.Popup = formInfo.Popup.Value; } catch { }
            try { if (formInfo.Modal.HasValue) frm.Modal = formInfo.Modal.Value; } catch { }
            try { if (formInfo.NavigationButtons.HasValue) frm.NavigationButtons = formInfo.NavigationButtons.Value; } catch { }
            try { if (formInfo.RecordSelectors.HasValue) frm.RecordSelectors = formInfo.RecordSelectors.Value; } catch { }
            try { if (formInfo.ScrollBars.HasValue) frm.ScrollBars = formInfo.ScrollBars.Value; } catch { }
            try { if (formInfo.BorderStyle.HasValue) frm.BorderStyle = formInfo.BorderStyle.Value; } catch { }
            try { if (formInfo.MinMaxButtons.HasValue) frm.MinMaxButtons = formInfo.MinMaxButtons.Value ? 3 : 0; } catch { }
            try { if (formInfo.Width.HasValue) frm.Width = formInfo.Width.Value; } catch { }

            if (formInfo.Controls != null && formInfo.Controls.Count > 0)
            {
                foreach (var ctrl in formInfo.Controls)
                {
                    int controlType = ctrl.Type?.ToLowerInvariant() switch
                    {
                        "label"         => acLabel,
                        "commandbutton" => acCommandButton,
                        "subform"       => acSubform,
                        "rectangle"     => acRectangle,
                        _               => acTextBox
                    };
                    dynamic newCtrl = _accessApp.CreateControl(tempName, controlType, acDetail,
                        System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                        ctrl.Left, ctrl.Top, ctrl.Width, ctrl.Height);
                    try { newCtrl.Name = ctrl.Name; } catch { }
                    try { if (!string.IsNullOrEmpty(ctrl.Caption)) newCtrl.Caption = ctrl.Caption; } catch { }
                    try { if (!string.IsNullOrEmpty(ctrl.ControlSource)) newCtrl.ControlSource = ctrl.ControlSource; } catch { }
                    try { if (!string.IsNullOrEmpty(ctrl.SourceObject)) newCtrl.SourceObject = ctrl.SourceObject; } catch { }
                    try { if (!string.IsNullOrEmpty(ctrl.LinkChildFields)) newCtrl.LinkChildFields = ctrl.LinkChildFields; } catch { }
                    try { if (!string.IsNullOrEmpty(ctrl.LinkMasterFields)) newCtrl.LinkMasterFields = ctrl.LinkMasterFields; } catch { }
                    try { newCtrl.Visible = ctrl.Visible; } catch { }
                    try { newCtrl.Enabled = ctrl.Enabled; } catch { }
                    try { if (ctrl.BackColor.HasValue) newCtrl.BackColor = ctrl.BackColor.Value; } catch { }
                    try { if (ctrl.ForeColor.HasValue) newCtrl.ForeColor = ctrl.ForeColor.Value; } catch { }
                    try { if (ctrl.FontBold.HasValue) newCtrl.FontBold = ctrl.FontBold.Value; } catch { }
                    try { if (ctrl.FontSize.HasValue) newCtrl.FontSize = ctrl.FontSize.Value; } catch { }
                }
            }
            else
            {
                dynamic placeholder = _accessApp.CreateControl(tempName, acTextBox, acDetail,
                    System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                    100, 100, 1000, 300);
                try { placeholder.Name = "PlaceholderControl"; } catch { }
            }

            // Save first with the auto-generated name to prevent Access showing a "Save As" dialog
            // (in German Access, DoCmd.Close with acSaveYes on an unsaved form triggers a prompt)
            try { _accessApp.DoCmd.Save(acForm, tempName); } catch { }
            _accessApp.DoCmd.Close(acForm, tempName, acSaveYes);
            _accessApp.DoCmd.Rename(formInfo.Name, acForm, tempName);
        }

        public void DeleteForm(string formName)
        {
            EnsureAccessApp();
            const int acForm = 2;
            _accessApp!.DoCmd.DeleteObject(acForm, formName);
        }

        public string ExportReportToText(string reportName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            EnsureAccessApp();
            const int acReport = 3;
            const int acDesign = 1;
            const int acSaveNo = 2;

            _accessApp!.DoCmd.OpenReport(reportName, acDesign);
            List<ControlInfo> controls;
            try
            {
                controls = ReadControlsFromCollection(_accessApp.Reports(reportName).Controls);
            }
            finally
            {
                try { _accessApp.DoCmd.Close(acReport, reportName, acSaveNo); } catch { }
            }

            var reportData = new
            {
                Name = reportName,
                ExportedAt = DateTime.UtcNow,
                Controls = controls
            };

            return JsonSerializer.Serialize(reportData, new JsonSerializerOptions { WriteIndented = true });
        }

        public void ImportReportFromText(string reportData)
        {
            EnsureAccessApp();

            var reportInfo = JsonSerializer.Deserialize<ReportExportData>(reportData);
            if (reportInfo == null) throw new ArgumentException("Invalid report data");
            if (string.IsNullOrWhiteSpace(reportInfo.Name)) throw new ArgumentException("Report name is required");

            const int acReport = 3;
            const int acSaveYes = 1;
            const int acDetail = 0;
            const int acTextBox = 109;
            const int acLabel = 100;
            const int acCommandButton = 104;
            const int acSubReport = 112;

            // Delete existing report if present
            try { _accessApp!.DoCmd.DeleteObject(acReport, reportInfo.Name); } catch { }

            dynamic rpt = _accessApp!.CreateReport();
            string tempName = (string)rpt.Name;

            if (reportInfo.Controls != null && reportInfo.Controls.Count > 0)
            {
                foreach (var ctrl in reportInfo.Controls)
                {
                    int controlType = ctrl.Type?.ToLowerInvariant() switch
                    {
                        "label"         => acLabel,
                        "commandbutton" => acCommandButton,
                        "subreport"     => acSubReport,
                        _               => acTextBox
                    };
                    dynamic newCtrl = _accessApp.CreateReportControl(tempName, controlType, acDetail,
                        System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                        ctrl.Left, ctrl.Top, ctrl.Width, ctrl.Height);
                    try { newCtrl.Name = ctrl.Name; } catch { }
                    try { if (!string.IsNullOrEmpty(ctrl.Caption)) newCtrl.Caption = ctrl.Caption; } catch { }
                    try { if (!string.IsNullOrEmpty(ctrl.ControlSource)) newCtrl.ControlSource = ctrl.ControlSource; } catch { }
                    try { newCtrl.Visible = ctrl.Visible; } catch { }
                    try { if (ctrl.BackColor.HasValue) newCtrl.BackColor = ctrl.BackColor.Value; } catch { }
                    try { if (ctrl.ForeColor.HasValue) newCtrl.ForeColor = ctrl.ForeColor.Value; } catch { }
                    try { if (ctrl.FontBold.HasValue) newCtrl.FontBold = ctrl.FontBold.Value; } catch { }
                    try { if (ctrl.FontSize.HasValue) newCtrl.FontSize = ctrl.FontSize.Value; } catch { }
                }
            }

            try { _accessApp.DoCmd.Save(acReport, tempName); } catch { }
            _accessApp.DoCmd.Close(acReport, tempName, acSaveYes);
            _accessApp.DoCmd.Rename(reportInfo.Name, acReport, tempName);
        }

        public void DeleteReport(string reportName)
        {
            EnsureAccessApp();
            const int acReport = 3;
            _accessApp!.DoCmd.DeleteObject(acReport, reportName);
        }

        #endregion

        #region Helper Methods

        private List<FieldInfo> GetTableFields(string tableName)
        {
            var fields = new List<FieldInfo>();
            
            try
            {
                var schema = _oleDbConnection!.GetSchema("Columns", new string[] { null!, null!, tableName });
                
                foreach (System.Data.DataRow row in schema.Rows)
                {
                    fields.Add(new FieldInfo
                    {
                        Name = row["COLUMN_NAME"]?.ToString() ?? "",
                        Type = row["DATA_TYPE"]?.ToString() ?? "",
                        Size = Convert.ToInt32(row["CHARACTER_MAXIMUM_LENGTH"] ?? 0),
                        Required = row["IS_NULLABLE"]?.ToString() == "NO",
                        AllowZeroLength = true // Default value
                    });
                }
            }
            catch
            {
                // Return empty list if table doesn't exist or can't be accessed
            }

            return fields;
        }

        private long GetTableRecordCount(string tableName)
        {
            try
            {
                var command = new OleDbCommand($"SELECT COUNT(*) FROM [{tableName}]", _oleDbConnection);
                return Convert.ToInt64(command.ExecuteScalar());
            }
            catch
            {
                return 0;
            }
        }

        #endregion

        public void Dispose()
        {
            if (!_disposed)
            {
                Disconnect();
                _disposed = true;
            }
        }
    }

    #region Data Models

    public class CompileResult
    {
        public bool IsCompiled { get; set; }
        public List<string> Errors { get; set; } = new();
        /// <summary>Name of the module containing the first compile error (if available).</summary>
        public string? ErrorModule { get; set; }
        /// <summary>Approximate line number of the first compile error within the module (if available).</summary>
        public int? ErrorLine { get; set; }
    }

    public class TableInfo
    {
        public string Name { get; set; } = "";
        public List<FieldInfo> Fields { get; set; } = new();
        public long RecordCount { get; set; }
    }

    public class FieldInfo
    {
        public string Name { get; set; } = "";
        public string Type { get; set; } = "";
        public int Size { get; set; }
        public bool Required { get; set; }
        public bool AllowZeroLength { get; set; }
    }

    public class QueryInfo
    {
        public string Name { get; set; } = "";
        public string SQL { get; set; } = "";
        public string Type { get; set; } = "";
    }

    public class RelationshipInfo
    {
        public string Name { get; set; } = "";
        public string Table { get; set; } = "";
        public string ForeignTable { get; set; } = "";
        public string Attributes { get; set; } = "";
    }

    public class FormInfo
    {
        public string Name { get; set; } = "";
        public string FullName { get; set; } = "";
        public string Type { get; set; } = "";
    }

    public class ReportInfo
    {
        public string Name { get; set; } = "";
        public string FullName { get; set; } = "";
        public string Type { get; set; } = "";
    }

    public class MacroInfo
    {
        public string Name { get; set; } = "";
        public string FullName { get; set; } = "";
        public string Type { get; set; } = "";
    }

    public class ModuleInfo
    {
        public string Name { get; set; } = "";
        public string FullName { get; set; } = "";
        public string Type { get; set; } = "";
    }

    public class VBAProjectInfo
    {
        public string Name { get; set; } = "";
        public string Description { get; set; } = "";
        public List<VBAModuleInfo> Modules { get; set; } = new();
    }

    public class VBAModuleInfo
    {
        public string Name { get; set; } = "";
        public string Type { get; set; } = "";
        public bool HasCode { get; set; }
    }

    public class SystemTableInfo
    {
        public string Name { get; set; } = "";
        public DateTime DateCreated { get; set; }
        public DateTime LastUpdated { get; set; }
        public long RecordCount { get; set; }
    }

    public class MetadataInfo
    {
        public string Name { get; set; } = "";
        public string Type { get; set; } = "";
        public string Flags { get; set; } = "";
        public string DateCreated { get; set; } = "";
        public string DateModified { get; set; } = "";
    }

    public class ControlInfo
    {
        public string Name { get; set; } = "";
        public string Type { get; set; } = "";
        public int Left { get; set; }
        public int Top { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public bool Visible { get; set; } = true;
        public bool Enabled { get; set; } = true;
        public string? Caption { get; set; }
        public string? ControlSource { get; set; }
        public string? SourceObject { get; set; }
        public string? LinkChildFields { get; set; }
        public string? LinkMasterFields { get; set; }
        public int? BackColor { get; set; }
        public int? ForeColor { get; set; }
        public bool? FontBold { get; set; }
        public int? FontSize { get; set; }
        public bool? SpecialEffect { get; set; }
    }

    public class ControlProperties
    {
        public string Name { get; set; } = "";
        public string Type { get; set; } = "";
        public int Left { get; set; }
        public int Top { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public bool Visible { get; set; }
        public bool Enabled { get; set; }
        public int BackColor { get; set; }
        public int ForeColor { get; set; }
        public string FontName { get; set; } = "";
        public int FontSize { get; set; }
        public bool FontBold { get; set; }
        public bool FontItalic { get; set; }
    }

    public class FormExportData
    {
        public string Name { get; set; } = "";
        public DateTime ExportedAt { get; set; }
        public List<ControlInfo> Controls { get; set; } = new();
        public string VBA { get; set; } = "";
        public string? RecordSource { get; set; }
        public int? DefaultView { get; set; }  // 0=Single, 1=Continuous, 2=Datasheet, 3=PivotTable
        public bool? Popup { get; set; }
        public bool? Modal { get; set; }
        public bool? NavigationButtons { get; set; }
        public bool? RecordSelectors { get; set; }
        public int? ScrollBars { get; set; }  // 0=None, 1=Horiz, 2=Vert, 3=Both
        public int? BorderStyle { get; set; } // 0=None, 1=Thin, 2=Sizable, 3=Dialog
        public bool? MinMaxButtons { get; set; }
        public int? Width { get; set; }
    }

    public class ReportExportData
    {
        public string Name { get; set; } = "";
        public DateTime ExportedAt { get; set; }
        public List<ControlInfo> Controls { get; set; } = new();
    }

    #endregion
} 