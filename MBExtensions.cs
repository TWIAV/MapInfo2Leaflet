// Namespace Declaration
using System;
using System.Windows;
using System.Windows.Forms;
using System.IO;
using System.Text;

// The MBExtensions Namespace
namespace MBExtensions
{
   // The MBDialogs class
   class MBDialogs
	{
		public static void NoteAgain(string sNotification, string sCaption)
		{
		string notification = sNotification;
		string caption = sCaption;
		MessageBox.Show(notification, caption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
		}
	  public static bool AskAgain(string sQuestion, string sCaption)
		{
		string question = sQuestion;
		string caption = sCaption;
		var result = MessageBox.Show(question, caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
			if (result == DialogResult.Yes)
				{
					return true;
				}
			else
				{
					return false;
				}
		}
	  public static bool Warn(string sWarning, string sCaption)
		{
		string warning = sWarning;
		string caption = sCaption;
		var result = MessageBox.Show(warning, caption, MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);
			if (result == DialogResult.OK)
				{
					return true;
				}
			else
				{
					return false;
				}
		}
		public static string BrowseForFolder(string sDescription, string sFolder)
			{
				FolderBrowserDialog dlg = new FolderBrowserDialog();
				dlg.SelectedPath = sFolder;
				dlg.Description = sDescription;
				dlg.ShowNewFolderButton = true;
				if (dlg.ShowDialog() == DialogResult.OK)
					return dlg.SelectedPath;
				else
					return "";
			}

	}
	// The MBFileManagement class
	class MBFileManagement
	{
		public static void OpenFileOrDir(string sOpen)
			{
				System.Diagnostics.Process.Start(sOpen);
			}

		public static void CreateDir(string sDir)
			{
				bool exists = System.IO.Directory.Exists(sDir);
				if(!exists)
					System.IO.Directory.CreateDirectory(sDir);
			}
	}
	// The MBFileConversion class
	class MBFileConversion
	{
		public static void ConvertFileANSItoUTF8(string sFileOld, string sFileNew)
			{
				string contents = File.ReadAllText(sFileOld, Encoding.Default);
				File.WriteAllText(sFileNew, contents, Encoding.UTF8);
			}
	}
	// The MBCopyDir class
	class MBCopyDir
	{
		public static void CopyDir(string sSourceDirectory, string sTargetDirectory)
		{
			DirectoryInfo diSource = new DirectoryInfo(sSourceDirectory);
			DirectoryInfo diTarget = new DirectoryInfo(sTargetDirectory);

			CopyAll(diSource, diTarget);
		}

		public static void CopyAll(DirectoryInfo source, DirectoryInfo target)
		{
			// Check if the target directory exists; if not, create it.
			if (Directory.Exists(target.FullName) == false)
			{
				Directory.CreateDirectory(target.FullName);
			}

			// Copy each file into the new directory.
			foreach (FileInfo fi in source.GetFiles())
			{
				Console.WriteLine(@"Copying {0}\{1}", target.FullName, fi.Name);
				fi.CopyTo(Path.Combine(target.FullName, fi.Name), true);
			}

			// Copy each subdirectory using recursion.
			foreach (DirectoryInfo diSourceSubDir in source.GetDirectories())
			{
				DirectoryInfo nextTargetSubDir =
					target.CreateSubdirectory(diSourceSubDir.Name);
				CopyAll(diSourceSubDir, nextTargetSubDir);
			}
		}
	}
	// The MBDateAndTime class
	class MBDateAndTime
	{
      public static string RegionalLongDate(string sDateString)
      {
        string mbDateString;
        DateTime stringToDate;
 
        stringToDate = ConvertMBDate(sDateString);
 
        /* Convert the DateTime object stringToDate to
        it's equivalent long date string representation
        (depends on the Regional and Language Options) */
        mbDateString = (stringToDate.ToLongDateString());
         
        return mbDateString;
      }
       
      public static string RegionalMonth(string sDateString)
      {
        string mbMonthString;
        DateTime stringToDate;
 
        stringToDate = ConvertMBDate(sDateString);
 
        /* Extract from this date the name of the month
        in the regional language (depends on the Regional
        and Language Options) */
        mbMonthString = stringToDate.ToString("MMMM");
         
        return mbMonthString;
      }
 
      public static string RegionalWeekday(string sDateString)
      {
        string mbWeekdayString;
        DateTime stringToDate;
 
        stringToDate = ConvertMBDate(sDateString);
 
        /* Extract from this date the name of the weekday
        in the regional language (depends on the Regional
        and Language Options) */
        mbWeekdayString = stringToDate.ToString("dddd");
 
        return mbWeekdayString;
      }
 
      public static DateTime ConvertMBDate(string sDateString)
      /* The input string sDateString which comes from your
      MapBasic app will be either a Date (i.e. YYYYMMDD) or
      a DateTime (i.e. YYYYMMDDHHMMSSFFF) */
      {
          int mbYear;
          int mbMonth;
          int mbDay;
          DateTime stringToDate;
 
          /* Take the Year component from the input string
          (i.e. YYYY) and convert it to an integer */
          mbYear = Convert.ToInt32(sDateString.Substring(0, 4));
 
          /* Take the Month component from the input string
          (i.e. MM) and convert it to an integer */
          mbMonth = Convert.ToInt32(sDateString.Substring(4, 2));
 
          /* Take the Day component from the input string
          (i.e. DD) and convert it to an integer */
          mbDay = Convert.ToInt32(sDateString.Substring(6, 2));
 
          // Assign a value to the DateTime object stringToDate
          stringToDate = new DateTime(mbYear, mbMonth, mbDay);
 
          return stringToDate;
      }
  }
}
