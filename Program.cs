using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;

namespace ListFieldMigrationRecursive
{
    class Program
    {
        //prompt user for the old field name to copy data from
        public static string GetOldFieldName()
        {
            Console.WriteLine("Enter the Display Name of the old field to copy from:");
            return Console.ReadLine();
        }
        //prompyt user for the new field to copy data into
        public static string GetNewFieldName()
        {
            Console.WriteLine("");
            Console.WriteLine("Enter the Display Name of the new field to copy to:");
            return Console.ReadLine();
            
        }
        //create a list of all old/new field pairs for copying
        public static List<KeyValuePair<string, string>> GetFieldsToCopy()
        {
            List<KeyValuePair<string, string>> fieldsList = new List<KeyValuePair<string, string>>();
            bool isDone = false;
            while (isDone == false)
            {
                //prompt the user for the field names, add them to the list
                KeyValuePair<string, string> newPair = new KeyValuePair<string, string>(GetOldFieldName(), GetNewFieldName());
                fieldsList.Add(newPair);
                Console.WriteLine("");
                //propmt the user if they want to add more fields
                Console.WriteLine("Would you like to copy another field?(y/n)");
                string done = Console.ReadLine().ToLower().Trim();
                //if no more fields, set "isDone" to true, which will end the while loop
                if (done == "n" || done == "no")
                {
                    isDone = true;
                }
                else
                {
                    Console.WriteLine("");
                }
            }
            return fieldsList;
        }
        //method migrates data for the SPWeb object passed as the first argument
        //then calls itself for all that SPWeb's child SPWeb objects
        public static void MigrateDataRecursive(SPWeb oWeb, string listname, List<KeyValuePair<string, string>> copyFields, bool delete)
        {
            bool listExists = true;
            SPList oList = null;
            //try to get the list. if it fails, set the boolean to false
            try
            {
                oList = oWeb.Lists[listname];
            }
            catch
            {
                listExists = false;
            }
            if (listExists)
            {
                //save state of list properties that need to be changed so they can be restored later
                bool isForcedCheckout = oList.ForceCheckout;
                bool isEnabledModeration = oList.EnableModeration;
                oList.ForceCheckout = false;
                oList.EnableModeration = false;
                oList.Update();
                //create list of files with minor versions that need to be treated specially
                List<SPFile_Info> minorFileInfo = new List<SPFile_Info>();
                try
                {
                    foreach (SPFile oFile in oList.RootFolder.Files)
                    {
                        try
                        {
                            //if file is checked out, force a check in
                            if (oFile.CheckOutType != SPFile.SPCheckOutType.None)
                            {
                                oFile.CheckIn("");
                            }
                            //if file is a minor version or has no major version, add its relevent data to our list of minor version files
                            if (oFile.MinorVersion != 0 && oFile.MajorVersion != 0)
                            {
                                SPFile_Info currentFile = new SPFile_Info();
                                currentFile.majorVersion = oFile.MajorVersion;
                                currentFile.minorVersion = oFile.MinorVersion;
                                currentFile.Id = oFile.UniqueId;
                                currentFile.modifiedBy = oFile.Item["Modified By"];
                                currentFile.modifiedDate = oFile.Item["Modified"];
                                minorFileInfo.Add(currentFile);
                            }
                            //if the file is a major version or has no major version, migrate the column data and use SystemUpdate(false) to prevent any changes to versioning or modified data
                            else
                            {
                                foreach (KeyValuePair<string, string> oPair in copyFields)
                                {
                                    if (oFile.Item.Fields.ContainsField(oPair.Key))
                                    {
                                        if (oFile.Item[oPair.Key] != null)
                                        {
                                            oFile.Item[oPair.Value] = oFile.Item[oPair.Key].ToString();
                                        }
                                    }
                                }
                                oFile.Item.SystemUpdate(false);
                            }
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e);
                        }
                    }

                    //restore the most recent major version of all files in our minor version list, and publish it to create a new major version
                    foreach (SPFile_Info oFile_Info in minorFileInfo)
                    {
                        try
                        {
                            SPFile oFile = oWeb.GetFile(oFile_Info.Id);
                            oFile.CheckOut();
                            oFile.Versions.RestoreByLabel(oFile_Info.majorVersion + ".0");
                            oFile.CheckIn("");
                            oFile.Publish("");
                            oFile.Update();
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e);
                        }
                    }

                    //migrate the column data on the newly restored major version and use SystemUpdate(false) to prevent any changes to versioning or modified data
                    foreach (SPFile_Info oFile_Info in minorFileInfo)
                    {
                        try
                        {
                            SPListItem oItem = oWeb.GetFile(oFile_Info.Id).Item;
                            foreach (KeyValuePair<string, string> oPair in copyFields)
                            {
                                if (oItem.Fields.ContainsField(oPair.Key))
                                {
                                    if (oItem[oPair.Key] != null)
                                    {
                                        oItem[oPair.Value] = oItem[oPair.Key].ToString();
                                    }
                                }
                            }
                            oItem["Modified By"] = oFile_Info.modifiedBy.ToString();
                            oItem["Modified"] = oFile_Info.modifiedDate.ToString();
                            oItem.SystemUpdate(false);
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e);
                        }
                    }

                    //restore the original minor version of each file, and check it in to create a new minor version
                    foreach (SPFile_Info oFile_Info in minorFileInfo)
                    {
                        try
                        {
                            SPFile oFile = oWeb.GetFile(oFile_Info.Id);
                            oFile.CheckOut();
                            oFile.Versions.RestoreByLabel(oFile_Info.majorVersion + "." + oFile_Info.minorVersion);
                            oFile.CheckIn("");
                            oFile.Update();
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e);
                        }
                    }

                    //migrate the column data on the newly restored minor version and use SystemUpdate(false) to prevent any changes to versioning or modified data
                    foreach (SPFile_Info oFile_Info in minorFileInfo)
                    {
                        try
                        {
                            SPListItem oItem = oWeb.GetFile(oFile_Info.Id).Item;
                            foreach (KeyValuePair<string, string> oPair in copyFields)
                            {
                                if (oItem.Fields.ContainsField(oPair.Key))
                                {
                                    if (oItem[oPair.Key] != null)
                                    {
                                        oItem[oPair.Value] = oItem[oPair.Key].ToString();
                                    }
                                }
                            }
                            oItem["Modified By"] = oFile_Info.modifiedBy.ToString();
                            oItem["Modified"] = oFile_Info.modifiedDate.ToString();
                            oItem.SystemUpdate(false);
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e);
                        }
                    }
                    //if user selected to remove old columns from list, remove them
                    if (delete)
                    {
                        foreach (KeyValuePair<string, string> oPair in copyFields)
                        {
                            if (oList.Fields.ContainsField(oPair.Key))
                            {
                                oList.Fields[oPair.Key].Delete();
                            }
                        }
                        oList.Update();
                    }
                }
                finally
                {
                    //restore original list property values
                    oList.EnableModeration = isEnabledModeration;
                    oList.ForceCheckout = isForcedCheckout;
                    oList.Update();
                }
                Console.WriteLine(oWeb.ServerRelativeUrl + " Complete.");
            }
            //if the list does not exist
            else
            {
                Console.WriteLine("List does not exist in this subsite");
            }
            //call migration method on all child SPWebs using the same arguments passed to this one
            foreach (SPWeb newWeb in oWeb.Webs)
            {
                MigrateDataRecursive(newWeb, listname, copyFields, delete);
            }
        }
        static void Main(string[] args)
        {
            # region Set SPSite
            //user enters site URL in console window
            Console.WriteLine("Enter the URL of the target site:");
            SPSite oSite = new SPSite(Console.ReadLine());
            Console.WriteLine("");

            //specify a site URL instead, use for debugging purposes
            //SPSite oSite = new SPSite(@"http://yourserver");
            # endregion

            #region Set ListName
            Console.WriteLine("What is the name of the libraries on which you want to perform the migrate action?");
            string ListName = Console.ReadLine();

            //specify the name of the lists/libraries to migrate, use for debugging purposes
            //string ListName = "Pages";
            #endregion

            # region Set CopyFields
            //user enters pairs of fields for migration
            List<KeyValuePair<string, string>> CopyFields = GetFieldsToCopy();

            //specify pairs of fields for migration instead, use this for debugging purposes
            //List<KeyValuePair<string, string>> CopyFields = new List<KeyValuePair<string, string>>();
            //CopyFields.Add(new KeyValuePair<string,string>("OldColumn1", "NewColumn1"));
            //CopyFields.Add(new KeyValuePair<string,string>("Oldcolumn2", "NewColumn2"));
            # endregion

            SPWeb rootWeb = oSite.RootWeb;
            bool deleteOld = false;

            //propmt user if they want to remove old columns from lists
            Console.WriteLine("Do you wish to delete the old columns from the lists once the migration action is complete?(y/n)");
            string deleteFields = Console.ReadLine().ToLower().Trim();
            if (deleteFields == "y" || deleteFields == "yes")
            {
                deleteOld = true;
            }
            //call the migration method on the root subsite
            try
            {
                MigrateDataRecursive(rootWeb, ListName, CopyFields, deleteOld);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            //if the user opted to remove the old columns, check if they need to remove them as site columns as well.
            if (deleteOld)
            {
                Console.WriteLine("");
                Console.WriteLine("Do you wish to delete the Site Columns cooresponding to the old list columns?(y/n)");
                string deleteSiteColumns = Console.ReadLine().ToLower().Trim();
                //if yes, remove the site columns
                if (deleteSiteColumns == "y" || deleteSiteColumns == "yes")
                {
                    foreach (KeyValuePair<string, string> oPair in CopyFields)
                    {
                        try
                        {
                            SPField oField = rootWeb.Fields[oPair.Key];
                            //first loop through all content types and remove the site columns so they can be deleted
                            foreach (SPContentType oContentType in rootWeb.ContentTypes)
                            {
                                try
                                {
                                    if (oContentType.Fields.Contains(oField.Id))
                                    {
                                        oContentType.FieldLinks.Delete(oField.Id);
                                        oContentType.Update();
                                    }
                                }
                                catch (Exception e)
                                {
                                    Console.WriteLine(e);
                                }
                            }
                            //delete the fields
                            oField.AllowDeletion = true;
                            oField.Delete();
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e);
                        }
                    }
                }
            }
            Console.WriteLine("");
            Console.WriteLine("Operation Complete");
            Console.ReadKey();
        }
    }
}
