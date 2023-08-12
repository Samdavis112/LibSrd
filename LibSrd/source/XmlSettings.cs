using System;
using System.IO;
using System.Xml.Serialization;

// settings management class. Will need to copy actual class, as the code needs modifying for each program this is used with. (If not using normal Application settings such as for click once.)
namespace LibSrd
{
    /// <summary>
    /// A simple class to manage program settings using an Xml file. 
    /// Load with XmlSettings mySettings = XmlSettings.Read(@"C:\Settings.xml");  [NB static!]
    /// Save with mySettings.Save(@"C:\Settings.xml");                          [Not static]
    /// * Cannot use Dictionaries or other non-serialisable collections in this class *
    /// \\GBRNE-EU-SQL02\Datawarehouse\UploadControl\Settings.xml
    /// </summary>
    public class XmlSettings
    {
        // Settings go here:
        public static string SettingsFile = @"\\GBRNE-EU-SQL02\Datawarehouse\UploadControl\Settings.xml";


        #region MySettings stuff
        public string ErrMsg = null;

        /// <summary>
        /// Reading in settings has to be done using a static class as you cant deserialise an instance into itself.
        /// So use XmlSettings mySettings = XmlSettings.Read(@"C:\Settings.xml");
        /// (Options would be to use an external Helper class or use a Settings struct rather than an class).
        /// By default the class default values are written out if no settings file is found. The resulting xml file may then 
        /// be edited as required. 
        /// </summary>
        /// <param name="xmlFilename">e.g. Settings.xml. If null then defaults to XmlSettings.SettingsFile static class variable.</param>
        /// <param name="WriteDefaultIfNoSettingsFileFound">Set true to write out the default values if no settings file found.</param>
        /// <returns></returns>
        public static XmlSettings Read(string xmlFilename = null, bool WriteDefaultIfNoSettingsFileFound = true)
        {
            if (xmlFilename == null) xmlFilename = XmlSettings.SettingsFile;


            if (File.Exists(xmlFilename))
            {
                XmlSerializer deserializer = new XmlSerializer(typeof(XmlSettings));
                using (TextReader textReader = new StreamReader(xmlFilename))
                {
                    return (XmlSettings)deserializer.Deserialize(textReader);
                }
            }
            else
            {
                XmlSettings tmp = new XmlSettings();
                string err = null;
                if (WriteDefaultIfNoSettingsFileFound) err = tmp.Save(xmlFilename);
                tmp.ErrMsg = err;
                return tmp;
            }
        }

        /// <summary>
        /// Saves the settings to a file for next time.
        /// </summary>
        /// <param name="xmlFilename">e.g. Settings.xml. If null then defaults to XmlSettings.SettingsFile static class variable.</param>
        /// <param name="Backup">Set true to back up and existing settings file to [xmlFilename]%) </param>
        /// <returns>Error message or null</returns>
        public string Save(string xmlFilename = null, bool Backup = true)
        {
            if (xmlFilename == null) xmlFilename = XmlSettings.SettingsFile;

            string folder = Path.GetDirectoryName(xmlFilename);
            if (!Directory.Exists(folder))
            {
                return "Error in MySettings.Save - cannot access folder: '" + folder + "'";
            }

            if (Backup)
            {
                string backup = xmlFilename + "%";
                if (File.Exists(backup)) File.Delete(backup);
                if (File.Exists(xmlFilename)) File.Move(xmlFilename, backup);
            }
            XmlSerializer serializer = new XmlSerializer(typeof(XmlSettings));
            string msg = "";
            for (int i = 0; i < 10; i++)
            {
                try
                {
                    TextWriter textWriter = new StreamWriter(xmlFilename);
                    serializer.Serialize(textWriter, this);
                    textWriter.Close();
                    return null;
                }
                catch (Exception ex)
                {
                    msg = ex.Message;
                    System.Threading.Thread.Sleep(1000);
                }
            }
            throw new Exception("MySettings cannot save: " + msg);
        }
        #endregion
    }
}