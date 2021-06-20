using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using UnityEditor;
using UnityEngine;

public class UsingCharList : EditorWindow
{
    static string[]        exts       = new string[] { ".xls", ".txt", ".log", ".cs" };

    const string           ITEM_NAME  =  "Tools/Using Character List"; // コマンド名
    static readonly string CLASS_NAME = $"{nameof(UsingCharList)}";    // ウィンドウのタイトル

    [System.Serializable]
    public class DataParameter
    {
        /// <summary></summary>
        public bool     Enabled;
        /// <summary></summary>
        public string   Filename;
        /// <summary></summary>
        public string   ImportPath;
        
        /// <summary></summary>
        public DataParameter(bool enabled, string characterTextName, string importPath)
        {
            Enabled    = enabled;
            Filename   = characterTextName;
            ImportPath = importPath;
        }
    }

    class Entity
    {
        public HashSet<string>  Chars;
        public string           Name;

        public Entity(string name)
        {
            Name    = name;
            Chars   = new HashSet<string>();
        }
    }

    [System.Serializable]
    public class JsonParameter
    {
        public List<DataParameter>	Data = new List<DataParameter>();
    }

    static Vector2			curretScroll  = Vector2.zero;
    static string			jsonfile      = null;
    static JsonParameter	defines;

    static bool[]           checkExts = new bool[exts.Length];
    static List<string>     enableExts;

    /// <summary>
    /// ウィンドウを開きます
    /// </summary>
    [MenuItem(ITEM_NAME)]
    static void MenuExec()
    {
        var window = GetWindow<UsingCharList>(true, CLASS_NAME);
        window.Init();
    }

    /// <summary>
    /// ウィンドウオープンの可否を取得します
    /// </summary>
    [MenuItem(ITEM_NAME, true)]
    static bool CanCreate()
    {
        bool enable = !EditorApplication.isPlaying && !Application.isPlaying && !EditorApplication.isCompiling;
        if (enable == false)
        {
            Debug.Log ($"{CLASS_NAME}: can't create. wait seconds.");
        }
        return enable;
    }

    const string PREFS_USING_CHARLIST_ENABLE_EXT = "PREFS_USING_CHARLIST_ENABLE_EXT";
    /// <summary>
    /// 初期化
    /// </summary>
    void Init()
    {
        for (int i = 0; i < checkExts.Length; i++)
        {
            checkExts[i] = EditorPrefs.GetBool($"{PREFS_USING_CHARLIST_ENABLE_EXT}{i}", true);
        }

        // クラス名を取得
        string className = this.ToString().Trim().Replace("(", "").Replace(")", "");
        
        // クラス名.cs というファイルを検索
        string[]	files = System.IO.Directory.GetFiles(Application.dataPath, className + ".cs", System.IO.SearchOption.AllDirectories);
        string		file  = null;
        if (files.Length == 1)
        {
            // 同じパスに クラス名.json としてデータを格納
            file = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(files[0]), className + ".json");
            
            jsonfile = file;
        }
        
        defines = null;
        
        // クラス名.json があればそのデータを読み込む
        if (System.IO.File.Exists(jsonfile) == true)
        {
            string data = null;
            
            using (System.IO.StreamReader sr = new System.IO.StreamReader(jsonfile))
            {
                data = sr.ReadToEnd();
            }
            defines = JsonUtility.FromJson<JsonParameter>(data);
        }

        if (defines == null)
        {
            defines = new JsonParameter();
        }
    }

    /// <summary>
    /// GUI を表示する時に呼び出されます
    /// </summary>
    void OnGUI()
    {
        if (defines == null)
        {
            Close();
            return;
        }
        
        GUILayout.Space(20);

        GUILayout.Label("enable import extensions", EditorStyles.boldLabel);
        GUILayout.Space(8);
        GUILayout.BeginHorizontal();
        for (int i = 0; i < checkExts.Length; i++)
        {
            checkExts[i] = EditorGUILayout.ToggleLeft(exts[i], checkExts[i], GUILayout.Width(60));
        }
        GUILayout.EndHorizontal();

        GUILayout.Space(20);
        
        // 設定項目
        curretScroll = EditorGUILayout.BeginScrollView(curretScroll);
        EditorGUILayout.BeginVertical("box");
        int cnt = 0;
        foreach (DataParameter def in defines.Data)
        {
            def.Enabled	   = EditorGUILayout.BeginToggleGroup($"group(No. {cnt})", def.Enabled);
            def.ImportPath = EditorGUILayout.TextField("import (file, directory): ", def.ImportPath);
            def.Filename   = EditorGUILayout.TextField("export (txt): ", def.Filename);
            GUILayout.Space(20);
            EditorGUILayout.EndToggleGroup();
            
            cnt++;
        }
        EditorGUILayout.EndVertical();
        EditorGUILayout.EndScrollView();
        
        // ＋－
        using (new EditorGUILayout.HorizontalScope())
        {
            if (GUILayout.Button("+"))
            {
                defines.Data.Add(new DataParameter(true, "", ""));
            }
            if (GUILayout.Button("-"))
            {
                defines.Data.RemoveAt(defines.Data.Count-1);
            }
        }
        
        GUILayout.Space(20);
        
        if (GUILayout.Button("Create"))
        {
            // 情報として json ファイルに格納
            if (string.IsNullOrEmpty(jsonfile) == false)
            {
                string data = JsonUtility.ToJson(defines);
                
                data = data.Replace("¥r¥n", "¥n");
                using (System.IO.StreamWriter sw = new System.IO.StreamWriter(jsonfile, false))
                {
                    sw.Write(data);
                }
            }
            
            // enable extensions
            for (int i = 0; i < checkExts.Length; i++)
            {
                EditorPrefs.SetBool($"{PREFS_USING_CHARLIST_ENABLE_EXT}{i}", checkExts[i]);
            }
            
            // char file を作成
            enableExts = new List<string>();
            for (int i = 0; i < checkExts.Length; i++)
            {
                if (checkExts[i] == true)
                {
                    enableExts.Add(exts[i]);
                }
            }

            string filelist = "";
            string crlf     = System.Environment.NewLine;
            
            int errcnt = 0;

            try
            {
                cnt = 0;
                foreach (var x in defines.Data)
                {
                    if (x.Enabled == false)
                    {
                        continue;
                    }

                    if (cancelableProgressBar(cnt, defines.Data.Count, x.Filename) == true)
                    {
                        return;
                    }

                    int length = create(x);
                    if (length > 0)
                    {
                        filelist += $"{Path.GetFileName(x.Filename)}({length}){crlf}";
                    }
                    else
                    {
                        filelist += $"{Path.GetFileName(x.Filename)} [FAILED]{crlf}";
                        errcnt++;
                    }
                    cnt++;
                }
            }
            finally
            {
                EditorUtility.ClearProgressBar();
            }
            
            // 強制的に unity database を更新
            AssetDatabase.Refresh(ImportAssetOptions.ForceUpdate);
            
            dialog($"{cnt} file(s) created. {errcnt} error(s)." + crlf + crlf + filelist);

            Close();
        }
    }
    
    /// <summary>
    /// キャラクターリストファイル作成
    /// </summary>
    static int create(DataParameter param)
    {
        int length = 0;
        
        try
        {
            List<string> files = new List<string>();
        
            var path = param.ImportPath;

            if (Directory.Exists(path) == true)
            {
                List<string> filesInDirectory =
                    new List<string>(Directory.GetFiles(path, "*", SearchOption.AllDirectories)).FindAll(
                        (a) =>
                        {
                            string ext = Path.GetExtension(a);
                            for (int i = 0; i < enableExts.Count; i++)
                            {
                                if (ext.IndexOf(enableExts[i]) >= 0)
                                {
                                    return true;
                                }
                            }
                            return false;
                        }
                    );
                files.AddRange(filesInDirectory);
            }
            else
            if (File.Exists(path) == true)
            {
                files.Add(path);
            }

            StringBuilder sb = new StringBuilder();

            foreach (string file in files)
            {
                var ext = Path.GetExtension(file).ToLower();

                Entity entity = new Entity(Path.GetFileNameWithoutExtension(path));

                if (ext == ".xls" || ext == ".xlsx")
                {
                    if (import_xls(file, entity.Chars) == false)
                    {
                        Debug.LogError($"import xls failed. [{file}]");
                        return 0;
                    }
                }
                else
                if (ext == ".txt")
                {
                    if (import_txt(file, entity.Chars) == false)
                    {
                        Debug.LogError($"import txt failed. [{file}]");
                        return 0;
                    }
                }

                foreach (var chara in entity.Chars)
                {
                    sb.Append(chara);
                }
            }

            var output         = sb.ToString();
            var outputFilename = param.Filename;

            length = sb.Length;

            completeDirectory(Path.GetDirectoryName(outputFilename));
            File.WriteAllText(outputFilename, output);
        }
        catch (Exception ex)
        {
            Debug.LogError($"{ex.Message}");
            length = 0;
        }
        
        return length;
    }

    /// <summary>
    /// 指定ディレクトリが存在しない場合、上から辿って作成する
    /// </summary>
    /// <param name="dir">指定ディレクトリ</param>
    /// <returns>true..作成した</returns>
    static bool completeDirectory(string dir)
    {
        if (string.IsNullOrEmpty(dir) == true)
        {
            return false;
        }
        if (Directory.Exists(dir) == false)
        {
            completeDirectory(Path.GetDirectoryName(dir));
            Directory.CreateDirectory(dir);
            return true;
        }
        
        return false;
    }

    /// <summary>
    /// テキストのキャラクターリスト作成
    /// </summary>
    static bool import_txt(string path, HashSet<string> chars)
    {
        bool succeed = true;

        try
        {
            string text = File.ReadAllText(path);

            for (int i = 0; i < text.Length; i++)
            {
                chars.Add(text[i].ToString());
            }
        }
        catch (Exception ex)
        {
            Debug.LogError($"{ex.Message}");
            succeed = false;
        }

        return succeed;
    }

    /// <summary>
    /// xlsx のキャラクターリスト作成
    /// </summary>
    static bool import_xls(string path, HashSet<string> chars)
    {
        bool succeed = true;

        using (FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        {
            IWorkbook book = null;
            if (Path.GetExtension(path) == ".xls")
            {
                book = new HSSFWorkbook(stream);
            }
            else
            {
                book = new XSSFWorkbook(stream);
            }

            for (int sheetno = 0; sheetno < book.NumberOfSheets; ++sheetno)
            {
                ISheet sheet = book.GetSheetAt(sheetno);

                int r = 0, c = 0;

                try
                {
                    for (r = 0; r <= sheet.LastRowNum; r++)
                    {
                        IRow row = sheet.GetRow(r);
                        if (row == null)
                        {
                            continue;
                        }
                        for (c = 0; c < row.Cells.Count; c++)
                        {
                            ICell    cell     = row.Cells[c];
                            string   cellstr  = cell.ToString();
                            CellType celltype = cell.CellType == CellType.Formula ? cell.CachedFormulaResultType : cell.CellType;

                            switch (celltype)
                            {
                                case CellType.Numeric:
                                    cellstr = cell.NumericCellValue.ToString();
                                    break;
                                case CellType.Boolean:
                                    cellstr = cell.BooleanCellValue.ToString();
                                    break;
                                case CellType.String:
                                    cellstr = cell.StringCellValue.ToString();
                                    break;
                            }

                            for (int i = 0; i < cellstr.Length; i++)
                            {
                                chars.Add(cellstr[i].ToString());
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.LogError($"{sheet.SheetName} R:{r} C:{c} {ex.Message}");
                    succeed = false;
                }
            }
        }

        return succeed;
    }

    /// <summary>
    /// ダイアログ表示
    /// </summary>
    static void dialog(string msg, params object[] objs)
    {
        if (objs != null)
        {
            msg = string.Format(msg, objs);
        }
        EditorUtility.DisplayDialog($"{CLASS_NAME}", msg, "ok");
    }

    /// <summary>
    /// エラー表示
    /// </summary>
    static void dialog_error(string msg, params object[] objs)
    {
        if (objs != null)
        {
            msg = string.Format(msg, objs);
        }
        EditorUtility.DisplayDialog($"{CLASS_NAME}", $"[ERROR]\r\n{msg}", "ok");
    }

    /// <summary>
    /// ok/cancel ダイアログ表示
    /// </summary>
    static bool dialog_select(string msg, params object[] objs)
    {
        if (objs != null)
        {
            msg = string.Format(msg, objs);
        }
        return EditorUtility.DisplayDialog($"{CLASS_NAME}", msg, "ok", "cancel");
    }

    /// <summary>
    /// キャンセルつき進捗バー (index+1)/max %
    /// </summary>
    static bool cancelableProgressBar(int index, int max, string msg)
    {
        float	perc = (float)(index+1) / (float)max;
        
        bool result =
            EditorUtility.DisplayCancelableProgressBar(
                $"{CLASS_NAME}",
                perc.ToString("00.0%") + "　" + msg,
                perc
            );
        if (result == true)
        {
            EditorUtility.ClearProgressBar();
            dialog("user cancel.");
            return true;
        }
        return false;
    }
}
