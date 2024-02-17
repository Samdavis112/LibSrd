using System;
using System.IO;

// settings management class. Will need to copy actual class, as the code needs modifying for each program this is used with. (If not using normal Application settings such as for click once.)
namespace LibSrd_NETCore
{
    /// <summary>
    /// A simple class to manage program settings using an Json file. 
    /// Load with JsonSettings mySettings = JsonSettings.Read(@"C:\Settings.json");  [NB static!]
    /// Save with mySettings.Save(@"C:\Settings.json");                          [Not static]
    /// * Cannot use Dictionaries or other non-serialisable collections in this class *
    /// </summary>
    public class JsonSettings
    {
        // Settings go here:
        public static string SettingsFile = @"C:\Users\sam07\OneDrive\Code\VisualStudio\repos\LibSrd\Tester\jsconfig1.json";


        #region JsonSettings stuff
        public string ErrMsg = null;

        /// <summary>
        /// Reading in settings has to be done using a static class as you cant deserialise an instance into itself.
        /// So use JsonSettings mySettings = JsonSettings.Read(@"C:\Settings.json");
        /// (Options would be to use an external Helper class or use a Settings struct rather than an class).
        /// By default the class default values are written out if no settings file is found. The resulting xml file may then 
        /// be edited as required. 
        /// </summary>
        /// <param name="Filepath">e.g. Settings.json. If null then defaults to JsonSettings.SettingsFile static class variable.</param>
        /// <param name="WriteDefaultIfNoSettingsFileFound">Set true to write out the default values if no settings file found.</param>
        /// <returns></returns>
        public static JsonSettings Read(string Filepath = null, bool WriteDefaultIfNoSettingsFileFound = true)
        {
            if (Filepath == null) Filepath = JsonSettings.SettingsFile;


            if (File.Exists(Filepath))
            {
                return (JsonSettings)Conversion.JsonToObject(File.ReadAllText(Filepath), typeof(JsonSettings));
            }
            else
            {
                JsonSettings tmp = new JsonSettings();
                string err = null;
                if (WriteDefaultIfNoSettingsFileFound) err = tmp.Save(Filepath);
                tmp.ErrMsg = err;
                return tmp;
            }
        }

        /// <summary>
        /// Saves the settings to a file for next time.
        /// </summary>
        /// <param name="Filepath">e.g. Settings.json. If null then defaults to JsonSettings.SettingsFile static class variable.</param>
        /// <param name="Backup">Set true to back up and existing settings file to [Filepath]%) </param>
        /// <returns>Error message or null</returns>
        public string Save(string Filepath = null, bool Backup = true)
        {
            if (Filepath == null) Filepath = JsonSettings.SettingsFile;

            string folder = Path.GetDirectoryName(Filepath);
            if (!Directory.Exists(folder))
            {
                return "Error in Settings.Save - cannot access folder: '" + folder + "'";
            }

            if (Backup)
            {
                string backup = Filepath + "%";
                if (File.Exists(backup)) File.Delete(backup);
                if (File.Exists(Filepath)) File.Move(Filepath, backup);
            }

            string msg = "";
            for (int i = 0; i < 10; i++)
            {
                try
                {
                    File.WriteAllText(Filepath, this.ToJson());
                    return null;
                }
                catch (Exception ex)
                {
                    msg = ex.Message;
                    System.Threading.Thread.Sleep(1000);
                }
            }
            throw new Exception("MySettings2 cannot save: " + msg);
        }
        #endregion
    }
}