/*=====================================================================
  
  This file is a Code Sample for Vault Extensions using Autodesk Vault API.

  Copyright (C) Autodesk Inc.  All rights reserved.

THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY
KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
PARTICULAR PURPOSE.
=====================================================================*/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

using Autodesk.Connectivity.Extensibility.Framework;
using Autodesk.Connectivity.Explorer.Extensibility;
using Autodesk.Connectivity.WebServices;
using Autodesk.Connectivity.WebServicesTools;
using Vault = Autodesk.DataManagement.Client.Framework.Vault;
using Autodesk.DataManagement.Client.Framework.Vault.Currency.Connections;
using Autodesk.DataManagement.Client.Framework.Vault.Currency.Entities;

[assembly: ApiVersion("10.0")]
[assembly: ExtensionId("056CFA08-1007-4918-B544-BC975E50EC2F")]

namespace ChangeMgmtExt
{
    public class mChangeMgmtEvents : IWebServiceExtension
    {
        public long mEcoId = new long();
        public string mCustEntUID = "";

        #region IWebServiceExtension Members
        public void OnLoad()
        {
            // File Events
            //DocumentService.AddFileEvents.GetRestrictions += new EventHandler<AddFileCommandEventArgs>(AddFileEvents_GetRestrictions);
            //DocumentService.CheckinFileEvents.GetRestrictions += new EventHandler<CheckinFileCommandEventArgs>(CheckinFileEvents_GetRestrictions);
            //DocumentService.CheckoutFileEvents.GetRestrictions += new EventHandler<CheckoutFileCommandEventArgs>(CheckoutFileEvents_GetRestrictions);
            //DocumentService.DeleteFileEvents.GetRestrictions += new EventHandler<DeleteFileCommandEventArgs>(DeleteFileEvents_GetRestrictions);
            //DocumentService.DownloadFileEvents.GetRestrictions += new EventHandler<DownloadFileCommandEventArgs>(DownloadFileEvents_GetRestrictions);
            DocumentServiceExtensions.UpdateFileLifecycleStateEvents.Pre += new EventHandler<UpdateFileLifeCycleStateCommandEventArgs>(UpdateFileLifecycleStateEvents_Pre);
            DocumentServiceExtensions.UpdateFileLifecycleStateEvents.GetRestrictions += new EventHandler<UpdateFileLifeCycleStateCommandEventArgs>(UpdateFileLifecycleStateEvents_GetRestrictions);
            DocumentServiceExtensions.UpdateFileLifecycleStateEvents.Post += new EventHandler<UpdateFileLifeCycleStateCommandEventArgs>(UpdateFileLifecycleStateEvents_Post);

            // Folder Events
            //DocumentService.AddFolderEvents.GetRestrictions += new EventHandler<AddFolderCommandEventArgs>(AddFolderEvents_GetRestrictions);
            //DocumentService.DeleteFolderEvents.GetRestrictions += new EventHandler<DeleteFolderCommandEventArgs>(DeleteFolderEvents_GetRestrictions);

            //// Item Events
            //ItemService.AddItemEvents.GetRestrictions += new EventHandler<AddItemCommandEventArgs>(AddItemEvents_GetRestrictions);
            //ItemService.CommitItemEvents.GetRestrictions += new EventHandler<CommitItemCommandEventArgs>(CommitItemEvents_GetRestrictions);
            //ItemService.ItemRollbackLifeCycleStatesEvents.GetRestrictions += new EventHandler<ItemRollbackLifeCycleStateCommandEventArgs>(ItemRollbackLifeCycleStatesEvents_GetRestrictions);
            //ItemService.DeleteItemEvents.GetRestrictions += new EventHandler<DeleteItemCommandEventArgs>(DeleteItemEvents_GetRestrictions);
            //ItemService.EditItemEvents.GetRestrictions += new EventHandler<EditItemCommandEventArgs>(EditItemEvents_GetRestrictions);
            //ItemService.PromoteItemEvents.GetRestrictions += new EventHandler<PromoteItemCommandEventArgs>(PromoteItemEvents_GetRestrictions);
            //ItemService.UpdateItemLifecycleStateEvents.GetRestrictions += new EventHandler<UpdateItemLifeCycleStateCommandEventArgs>(UpdateItemLifecycleStateEvents_GetRestrictions);

            //// Change Order Events
            //ChangeOrderService.AddChangeOrderEvents.GetRestrictions += new EventHandler<AddChangeOrderCommandEventArgs>(AddChangeOrderEvents_GetRestrictions);
            //ChangeOrderService.CommitChangeOrderEvents.GetRestrictions += new EventHandler<CommitChangeOrderCommandEventArgs>(CommitChangeOrderEvents_GetRestrictions);
            //ChangeOrderService.DeleteChangeOrderEvents.GetRestrictions += new EventHandler<DeleteChangeOrderCommandEventArgs>(DeleteChangeOrderEvents_GetRestrictions);
            //ChangeOrderService.EditChangeOrderEvents.GetRestrictions += new EventHandler<EditChangeOrderCommandEventArgs>(EditChangeOrderEvents_GetRestrictions);
            //ChangeOrderService.UpdateChangeOrderLifecycleStateEvents.GetRestrictions += new EventHandler<UpdateChangeOrderLifeCycleStateCommandEventArgs>(UpdateChangeOrderLifecycleStateEvents_GetRestrictions);

            // Custom Entity Events
            CustomEntityService.UpdateCustomEntityLifecycleStateEvents.GetRestrictions += new EventHandler<UpdateCustomEntityLifeCycleStateCommandEventArgs>(UpdateCustomEntityLifecycleStateEvents_GetRestrictions);
            CustomEntityService.UpdateCustomEntityLifecycleStateEvents.Post += UpdateCustomEntityLifecycleStateEvents_Post;

        }

        private void UpdateCustomEntityLifecycleStateEvents_Post(object sender, UpdateCustomEntityLifeCycleStateCommandEventArgs e)
        {
            //create link from ECO to CO planned
        }

        private void UpdateCustomEntityLifecycleStateEvents_GetRestrictions(object sender, UpdateCustomEntityLifeCycleStateCommandEventArgs e)
        {
            try
            {
                IWebService service = sender as IWebService;
                if (service == null)
                    return;

                //we need an active connection to proceed
                WebServiceCredentials cred = new WebServiceCredentials(service);
                using (WebServiceManager mgr = new WebServiceManager(cred))
                {
                    long currentUserId = mgr.SecurityService.SecurityHeader.UserId;

                    Connection mConnection = new Connection(mgr, mgr.WebServiceCredentials.VaultName,
                        mgr.WebServiceCredentials.SecurityHeader.UserId, mgr.WebServiceCredentials.ServerIdentities.FileServer,
                        AuthenticationFlags.Standard);
                    if (mConnection == null) return;

                    //check category - if
                    CustEnt[] mCustEnts = mgr.CustomEntityService.GetCustomEntitiesByIds(e.CustomEntityIds);
                    CustEnt mCustEnt = mCustEnts[0];
                    Cat[] mCats = mgr.CategoryService.GetCategoriesByEntityClassId("CUSTENT", true);
                    Cat mCat = mCats.FirstOrDefault(n => n.Name == "Change Request");
                    if (mCustEnt.Cat.CatId != mCat.Id)
                    {
                        e.AddRestriction(new ExtensionRestriction("Change Request Lifecycle", "This lifecycle applies to Change Request Object Types only." +
                            Environment.NewLine + "Either the selected object is assigned mistakenly to this lifecycle or the object's category is mistakenly changed."));
                        return;
                    }

                    LfCycDef[] mAllLfCycDefs = mgr.LifeCycleService.GetAllLifeCycleDefinitions();
                    LfCycDef mLfCycleDef = mAllLfCycDefs.FirstOrDefault(n => n.DispName == "Change Request");

                    //no reason to proceed, if the current file does not belong to this lifecycle
                    if (mCustEnt.LfCyc.LfCycDefId != mLfCycleDef.Id) return;

                    //no reason to proceed, if the target lifecycle is not the defined one to create ECO
                    LfCycState mLfCycTargetState = mLfCycleDef.StateArray.FirstOrDefault(n => n.DispName == "Completed with ECO");
                    if (e.ToStateIds[0] != mLfCycTargetState.Id)
                    {
                        //e.AddRestriction(new ExtensionRestriction("Change Request Lifecycle", "Target State does not match configuration"));
                        return;
                    }

                    //we allow single selection only
                    if (e.CustomEntityIds.Length != 1)
                    {
                        e.AddRestriction(new ExtensionRestriction("Change Request Lifecycle", "Multiple selections will not process..." + Environment.NewLine +
                                                "Select single ECR to get linked ECO created."));
                        return;
                    }

                    //everything checked ?

                    //we are sure to have an appropriate change request entity and proceed
                    List<long> mCustEntsIDs = new List<long>();
                    foreach (var item in mCustEnts)
                    {
                        mCustEntsIDs.Add(item.Id);
                    }

                    Lnk[] mFileLinks = mConnection.WebServiceManager.DocumentService.GetLinksByParentIds(mCustEntsIDs.ToArray(), new string[] { "FILE" });
                    Lnk[] mItemLinks = mConnection.WebServiceManager.DocumentService.GetLinksByParentIds(mCustEntsIDs.ToArray(), new string[] { "ITEM" });
                    Lnk[] mCustEntLinks = mConnection.WebServiceManager.DocumentService.GetLinksByParentIds(mCustEntsIDs.ToArray(), new string[] { "CUSTENT" });

                    //Prep the data that we need for the ECO: CO Props -> ECO Title, Description, UDP "ECR"
                    #region PropertyPreparation
                    PropertyService mPropSrvc = mConnection.WebServiceManager.PropertyService;
                    PropDef[] mPropDefs = mPropSrvc.GetPropertyDefinitionsByEntityClassId("CUSTENT").Where(n => n.DispName == "Title" || n.DispName == "Description"
                        || n.DispName == "Ident Number" || n.DispName == "Email Recipient").ToArray();

                    Dictionary<string, long> mProperties = new Dictionary<string, long>();
                    int i = 0;
                    foreach (var item in mPropDefs)
                    {
                        mProperties.Add(mPropDefs[i].DispName, mPropDefs[i].Id);
                        i += 1;
                    }

                    long[] mPropDefIdsArray = new long[mProperties.Count()];
                    mPropDefIdsArray[0] = mProperties["Title"];
                    mPropDefIdsArray[1] = mProperties["Description"];
                    mPropDefIdsArray[2] = mProperties["Ident Number"];
                    mPropDefIdsArray[3] = mProperties["Email Recipient"];
                    #endregion

                    #region prepfiles
                    //Prep the files to get linked to the ECO
                    List<long> mFileMasterIds = new List<long>();
                    if (mFileLinks != null)
                    {
                        foreach (var item in mFileLinks)
                        {
                            File mFile = mConnection.WebServiceManager.DocumentService.GetFileById(item.ToEntId);
                            if (mFile.ControlledByChangeOrder == false)
                            {
                                mFileMasterIds.Add(mFile.MasterId);
                            }
                            else
                            {
                                e.AddRestriction(new ExtensionRestriction("Change Request Lifecycle", "At least one of the linked files is already linked to an ECO." + Environment.NewLine +
                                               "Therefore the change of state with ECO creation is being refused."));
                                return;
                            }
                        }
                    }
                    #endregion
                    #region prep items
                    //Prep the items to get linked to the ECO
                    List<long> mItemMasterIds = new List<long>();
                    if (mItemLinks != null)
                    {
                        foreach (var item in mItemLinks)
                        {
                            Item[] mItems = mConnection.WebServiceManager.ItemService.GetItemsByIds(new long[] { item.ToEntId });
                            if (mItems[0].ControlledByChangeOrder == false)
                            {
                                mItemMasterIds.Add(mItems[0].MasterId);
                            }
                            else
                            {
                                e.AddRestriction(new ExtensionRestriction("Change Request Lifecycle", "At least one of the linked items is already linked to an ECO." + Environment.NewLine +
                                               "Therefore the change of state with ECO creation is being refused."));
                                return;
                            }
                        }
                    }
                    #endregion

                    #region prep ECO UDPs
                    PropInst[] mCustEntProps = mPropSrvc.GetProperties("CUSTENT", new long[] { mCustEntsIDs[0] }, mPropDefIdsArray);
                    //the properties are filtered to Title and Subject, but not sorted; we need to grab the right ones...
                    string mEcoTitle = "";
                    var mSinglePropInst = mCustEntProps.Where(n => n.PropDefId == mProperties["Title"]);
                    var mPropInstList = mSinglePropInst.ToList();
                    if (mPropInstList[0].Val != null) mEcoTitle = mPropInstList[0].Val.ToString();

                    string mEcoDescr = "";
                    mSinglePropInst = mCustEntProps.Where(n => n.PropDefId == mProperties["Description"]);
                    mPropInstList = mSinglePropInst.ToList();
                    if (mPropInstList[0].Val != null) mEcoDescr = mPropInstList[0].Val.ToString();

                    string mEmailAdr = "";
                    mSinglePropInst = mCustEntProps.Where(n => n.PropDefId == mProperties["Email Recipient"]);
                    mPropInstList = mSinglePropInst.ToList();
                    if (mPropInstList[0].Val != null) mEmailAdr = mPropInstList[0].Val.ToString();

                    PropDef[] mEcoUdpDefs = mPropSrvc.GetPropertyDefinitionsByEntityClassId("ChangeOrder").Where(n => n.DispName == "Initiated by ECR").ToArray();
                    Dictionary<string, long> mEcoProperties = new Dictionary<string, long>();
                    i = 0;
                    foreach (var item in mEcoUdpDefs)
                    {
                        mEcoProperties.Add(mEcoUdpDefs[i].DispName, mEcoUdpDefs[i].Id);
                        i += 1;
                    }

                    long[] mEcoPropDefIdsArray = new long[mEcoProperties.Count()];
                    mEcoPropDefIdsArray[0] = mEcoProperties["Initiated by ECR"];
                    PropInst[] mEcoUDPs = mPropSrvc.GetProperties("ChangeOrder", new long[] { mCustEntsIDs[0] }, mEcoPropDefIdsArray);
                    //alternatively use UDPs unique number Goto //alternative 2:
                    #region alternative1
                    //mSinglePropInst = mCustEntProps.Where(n => n.PropDefId == mProperties["Ident Number"]);
                    //mPropInstList = mSinglePropInst.ToList();
                    //if (mPropInstList[0].Val != null) mEcoUDPs[0].Val = mPropInstList[0].Val.ToString();
                    #endregion
                    #region alternative2
                    mEcoUDPs[0].Val = mCustEnts[0].Num;
                    #endregion
                    #endregion

                    #region prep Emails
                    Email[] mAllEmails = new Email[1];
                    if (mEmailAdr != "")
                    {
                        Email mEmail = new Email();
                        mEmail.ToArray = new string[] { mEmailAdr };
                        //mEmail.Subject = "New ECO " + coNumber + " created"; //to be completed after ECO is created
                        //mEmail.Body = "The ECO " + coNumber + "is available now; you are supposed to complete and open it"; //to be completed after ECO is created
                        mAllEmails[0] = mEmail;
                    }
                    else mAllEmails = null;
                    #endregion

                    #region ECO execution
                    ChangeOrderService mEcoService = mConnection.WebServiceManager.ChangeOrderService;
                    Autodesk.Connectivity.WebServices.ChangeOrder mEcoNew = new Autodesk.Connectivity.WebServices.ChangeOrder(); //note this is the webservice.changeorder; there is also a wrapper in VDF;
                    mClsChangeMgmtExt mChangeMgmtClass = new mClsChangeMgmtExt();
                    mEcoNew = mChangeMgmtClass.mCreateECObyCustEnt(mConnection, mEcoService, mEcoTitle, mEcoDescr, mFileMasterIds.ToArray(), mItemMasterIds.ToArray(), mEcoUDPs, mAllEmails);
                    #endregion
                    if (mEcoNew == null)
                    {
                        e.AddRestriction(new ExtensionRestriction("Change Request Lifecycle", "The target status requires an ECO created. The creation of it failed" + Environment.NewLine +
                                               "Therefore the change of state is being refused."));
                        return;
                    }

                }
            } //end try
            catch
            {
                e.AddRestriction(new ExtensionRestriction("Change Request Lifecycle", "The target status requires an ECO created. The creation of it failed" + Environment.NewLine +
                                                "Therefore the change of state is being refused."));
                return;
            }
        }

        private void UpdateFileLifecycleStateEvents_Pre(object sender, UpdateFileLifeCycleStateCommandEventArgs e)
        {
        }
        private void UpdateFileLifecycleStateEvents_Post(object sender, UpdateFileLifeCycleStateCommandEventArgs e)
        {
        }

        private void UpdateFileLifecycleStateEvents_GetRestrictions(object sender, UpdateFileLifeCycleStateCommandEventArgs e)
        {
            
            try
            {
                IWebService service = sender as IWebService;
                if (service == null)
                    return;

                //we need an active connection to proceed
                WebServiceCredentials cred = new WebServiceCredentials(service);
                using (WebServiceManager mgr = new WebServiceManager(cred))
                {
                    long currentUserId = mgr.SecurityService.SecurityHeader.UserId;

                    Connection mConnection = new Connection(mgr, mgr.WebServiceCredentials.VaultName,
                        mgr.WebServiceCredentials.SecurityHeader.UserId, mgr.WebServiceCredentials.ServerIdentities.FileServer,
                        AuthenticationFlags.Standard);
                    if (mConnection == null) return;

                    File mFile = mgr.DocumentService.GetLatestFileByMasterId(e.FileMasterIds[0]);
                    if (mFile.ControlledByChangeOrder == true)
                    {
                        e.AddRestriction(new ExtensionRestriction("Change Request Lifecycle", "ECR Form is already linked to ECO" + Environment.NewLine +
                            "Only sinlge relationship is allowed for File-ECR"));
                        return;
                    }

                    //file based ECR require the file to be of category "Change Request"
                    if (mFile.Cat.CatName != "Change Request") return;

                    //filter lifecycle to "Change Request"
                    long mCurrentLfcID = mFile.FileLfCyc.LfCycDefId;
                    LfCycDef[] mAllLfCycDefs = mgr.LifeCycleService.GetAllLifeCycleDefinitions();
                    LfCycDef mLfCycleDef = mAllLfCycDefs.FirstOrDefault(n => n.DispName == "Change Request");

                    //no reason to proceed, if the current file does not belong to this lifecycle
                    if (mFile.FileLfCyc.LfCycDefId != mLfCycleDef.Id) return;

                    //no reason to proceed, if the target lifecycle is not the defined one to create ECO
                    LfCycState mLfCycTargetState = mLfCycleDef.StateArray.FirstOrDefault(n => n.DispName == "Completed with ECO");
                    
                    for (long i = 0;  i < e.FileMasterIds.Length; i++) //int i = 0; i < widgets.Length; i++)
                    {
                        if (e.ToStateIds[i] != mLfCycTargetState.Id) return;
                    }

                    if (e.FileMasterIds.Length != 1)
                    {
                        e.AddRestriction(new ExtensionRestriction("Change Request Lifecycle", "Multiple selections will not process..." + Environment.NewLine +
                                                        "Select single ECR to get linked ECO created."));
                        return;
                    }

                    //everything is set to create the ECO and link the file
                    ChangeOrderService mEcoService = mConnection.WebServiceManager.ChangeOrderService;
                    long[] mEcoFiles = new long[] { mFile.Id };
                    long[] mEcoFileMaster = new long[] { mFile.MasterId };

                    Autodesk.Connectivity.WebServices.ChangeOrder mEcoNew = new Autodesk.Connectivity.WebServices.ChangeOrder();
                    mClsChangeMgmtExt mChangeMgmtClass = new mClsChangeMgmtExt();

                    mEcoNew = mChangeMgmtClass.mCreateECObyFile(mConnection, mEcoService, mEcoFileMaster, mEcoFiles);
                    if (mEcoNew == null)
                    {
                        e.AddRestriction(new ExtensionRestriction("Change Request Lifecycle", "The target status requires an ECO created. The creation of it failed" + Environment.NewLine +
                                               "Therefore the change of state is being refused."));
                        return;
                    }
                }
            }
            catch
            {
                e.AddRestriction(new ExtensionRestriction("Change Request Lifecycle", "The target status requires an ECO created. The creation of it failed" + Environment.NewLine +
                                                "Therefore the change of state is being refused."));
                return;
            }
        }
        #endregion
    } //end event class

    public class mClsChangeMgmtExt : IExplorerExtension
    {
        public const string mECRSystemName = "395e6304-6a49-4363-842c-3b791bffda09";
        public IEnumerable<CommandSite> CommandSites()
        {
            List<CommandSite> sites = new List<CommandSite>();
            // Describe user history command item
            //

            CommandItem mGoToEco = new CommandItem("Command.TestCommand2", "Go To ECO");
            SelectionTypeId mECRSelectionType = new SelectionTypeId(mECRSystemName);
            mGoToEco.NavigationTypes = new SelectionTypeId[] { mECRSelectionType };
            mGoToEco.MultiSelectEnabled = false;
            mGoToEco.Execute += mGoToEco_Execute;
            CommandSite mEcrContextMenu = new CommandSite("Menu.ECRContextMenu", "Go To ECO");
            CommandSiteLocation mECRCommandSiteLoc = new CommandSiteLocation(mECRSelectionType.EntityClassSubType, CommandSiteLocationType.ContextMenu);
            mEcrContextMenu.Location = mECRCommandSiteLoc;
            mEcrContextMenu.DeployAsPulldownMenu = false;
            mEcrContextMenu.AddCommand(mGoToEco);
            sites.Add(mEcrContextMenu);

            CommandItem mGoToECR = new CommandItem("Command.EcoContextMenu", "Go To ECR");
            mGoToECR.NavigationTypes = new SelectionTypeId[] { SelectionTypeId.ChangeOrder };
            mGoToECR.MultiSelectEnabled = false;
            mGoToECR.Execute += mGoToECR_Execute;
            CommandSite mEcoContextMenu = new CommandSite("Menu.ECOContextMenu", "Go To ECR");
            mEcoContextMenu.Location = CommandSiteLocation.ChangeOrderContextMenu;
            mEcoContextMenu.DeployAsPulldownMenu = false;
            mEcoContextMenu.AddCommand(mGoToECR);
            sites.Add(mEcoContextMenu);

            return sites;
        }

        private void mGoToECR_Execute(object sender, CommandItemEventArgs e)
        {
            try
            {
                Connection connection = e.Context.Application.Connection;
                CustomEntityService mCustEntSrvc = connection.WebServiceManager.CustomEntityService;

                // The Context part of the event args tells us information about what is selected.
                // Run some checks to make sure that the selection is valid.
                if (e.Context.CurrentSelectionSet.Count() == 0)
                    throw new Exception("Nothing is selected");
                else if (e.Context.CurrentSelectionSet.Count() > 1) //Menu definition already excludes multiselection
                    throw new Exception("This function does not support multiple selections");
                else
                {
                    // we only have one item selected, which is the expected behavior
                    ISelection selection = e.Context.CurrentSelectionSet.First();

                    Autodesk.Connectivity.WebServices.ChangeOrder[] mCOs = connection.WebServiceManager.ChangeOrderService.GetChangeOrdersByIds(new long[] { selection.Id });

                    List<long> mCoIDs = new List<long>();
                    foreach (var item in mCOs)
                    {
                        mCoIDs.Add(item.Id);
                    }

                    Autodesk.Connectivity.WebServices.ChangeOrder mCO = mCOs[0];
                    string mCoUID = mCO.Num;

                    PropertyService mPropSrvc = connection.WebServiceManager.PropertyService;
                    PropDef[] mPropDefs = mPropSrvc.GetPropertyDefinitionsByEntityClassId("CO").Where(n => n.DispName == "Initiated by ECR").ToArray();
                    Dictionary<string, long> mProperties = new Dictionary<string, long>();
                    int i = 0;
                    foreach (var item in mPropDefs)
                    {
                        mProperties.Add(mPropDefs[i].DispName, mPropDefs[i].Id);
                        i += 1;
                    }

                    long[] mPropDefIdsArray = new long[mProperties.Count()];
                    mPropDefIdsArray[0] = mProperties["Initiated by ECR"];

                    PropInst[] mCustEntProps = mPropSrvc.GetProperties("ChangeOrder", new long[] { mCoIDs[0] }, mPropDefIdsArray);
                    //the properties are filtered to Title and Subject, but not sorted; we need to grab the right ones...
                    string mEcrUID = "";
                    var mSinglePropInst = mCustEntProps.Where(n => n.PropDefId == mProperties["Initiated by ECR"]);
                    var mPropInstList = mSinglePropInst.ToList();
                    if (mPropInstList[0].Val != null) mEcrUID = mPropInstList[0].Val.ToString();
                    else
                    {
                        MessageBox.Show("No linked ECR found!" + Environment.NewLine + "Check Property 'Initiated by ECR'; this must not be empty", "Change Mangement Extension", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    //linked ECR may be a file, lets search for it
                    DocumentService mDocSrvc = connection.WebServiceManager.DocumentService;
                    Autodesk.Connectivity.WebServices.Folder mEcrFolder = mDocSrvc.GetFolderByPath("$/Change Requests");
                    List<File> mAllSrchResults = new List<Autodesk.Connectivity.WebServices.File>();
                    SrchCond[] mSrchCond = new SrchCond[1];
                    SrchSort[] mSrchSort = null;
                    mSrchCond[0] = mCreateSrchCond(connection, "Name", mEcrUID, "AND", EntityClassIds.Files);
                    SrchStatus mSrchStatus = null;
                    string bookmark = string.Empty;
                    while (mSrchStatus == null || mAllSrchResults.Count < mSrchStatus.TotalHits)
                    {
                        File[] mFileSearchResult = mDocSrvc.FindFilesBySearchConditions(mSrchCond, mSrchSort, new long[] { mEcrFolder.Id }, true, true, ref bookmark, out mSrchStatus);
                        if (mFileSearchResult != null)
                            mAllSrchResults.AddRange(mFileSearchResult);
                        else break;
                    }
                    if (mAllSrchResults.Count() == 1)
                    {
                        FileIteration mFileIter = new FileIteration(connection, mAllSrchResults[0]);
                        Autodesk.Connectivity.WebServices.Folder mFileFolder = connection.WebServiceManager.DocumentService.GetFolderById(mFileIter.FolderId);
                        string mFullFileName = mFileFolder.FullName + "/" + mAllSrchResults[0].Name;
                        SelectionTypeId mECRSelectionType = Autodesk.Connectivity.Explorer.Extensibility.SelectionTypeId.File;
                        LocationContext mLocation = new Autodesk.Connectivity.Explorer.Extensibility.LocationContext(mECRSelectionType, mFullFileName);
                        e.Context.GoToLocation = mLocation;
                        return;
                    }
                    //linked ECR != File -> it's a custom object
                    Autodesk.Connectivity.WebServices.CustEnt[] results = mCustEntSrvc.FindCustomEntitiesByNumbers(new string[] { mEcrUID });

                    if (results.Count() == 1)
                    {
                        SelectionTypeId mECRSelectionType = new SelectionTypeId(mECRSystemName);
                        LocationContext mLocation = new Autodesk.Connectivity.Explorer.Extensibility.LocationContext(mECRSelectionType, results[0].Num);
                        e.Context.GoToLocation = mLocation;
                    }
                }
            }
            catch (Exception ex)
            {
                // If something goes wrong, we don't want the exception to bubble up to Vault Explorer.
                throw new Exception("Unhandled Error: " + ex.Message);
            }
        }


        private void mGoToEco_Execute(object sender, CommandItemEventArgs e)
        {
            try
            {
                Connection connection = e.Context.Application.Connection;
                ChangeOrderService mCoSrvc = connection.WebServiceManager.ChangeOrderService;

                // The Context part of the event args tells us information about what is selected.
                // Run some checks to make sure that the selection is valid.
                if (e.Context.CurrentSelectionSet.Count() == 0)
                    throw new Exception("Nothing is selected");
                else if (e.Context.CurrentSelectionSet.Count() > 1)
                    throw new Exception("This function does not support multiple selections");
                else
                {
                    // we only have one item selected, which is the expected behavior

                    ISelection selection = e.Context.CurrentSelectionSet.First();

                    CustEnt[] mCustEnts = connection.WebServiceManager.CustomEntityService.GetCustomEntitiesByIds(new long[] { selection.Id });

                    List<long> mCustEntsIDs = new List<long>();
                    foreach (var item in mCustEnts)
                    {
                        mCustEntsIDs.Add(item.Id);
                    }

                    CustEnt mCustEnt = mCustEnts[0];
                    string mCustEntUID = mCustEnt.Num;

                    LfCycDef[] mAllLfCycDefs = connection.WebServiceManager.LifeCycleService.GetAllLifeCycleDefinitions();
                    LfCycDef mLfCycleDef = mAllLfCycDefs.FirstOrDefault(n => n.DispName == "Change Request");

                    //no reason to proceed, if the current object does not belong to this lifecycle
                    //if (mCustEnt.LfCyc.LfCycDefId != mLfCycleDef.Id) return;

                    //no reason to proceed, if the current lifecycle is not the defined completed one
                    LfCycState mLfCycTargetState = mLfCycleDef.StateArray.FirstOrDefault(n => n.DispName == "Completed with ECO");
                    if (mCustEnt.LfCyc.LfCycStateId != mLfCycTargetState.Id || (mCustEnt.LfCyc.LfCycDefId != mLfCycleDef.Id))
                    {
                        MessageBox.Show("ECR cannot link to ECO!" + Environment.NewLine + "Is your ECR in the state required?:" + Environment.NewLine +
                            " " + Environment.NewLine +
                            "1) ECR Lifecylce = 'Change Request'?" + Environment.NewLine +
                            "2) ECR Lifecycle Status = 'Completed with ECO'?"
                            , "Change Mangement Extension", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    SrchCond[] mSrchCond = new SrchCond[1];
                    SrchSort[] mSrchSort = null;
                    mSrchCond[0] = mCreateSrchCond(connection, "Initiated by ECR", mCustEnt.Num, "AND", EntityClassIds.ChangeOrders);
                    SrchStatus mSrchStatus = null;
                    string bookmark = string.Empty;

                    List<Autodesk.Connectivity.WebServices.ChangeOrder> mAllSrchResults = new List<Autodesk.Connectivity.WebServices.ChangeOrder>();
                    while (mSrchStatus == null || mAllSrchResults.Count < mSrchStatus.TotalHits)
                    {
                        Autodesk.Connectivity.WebServices.ChangeOrder[] results = mCoSrvc.FindChangeOrdersBySearchConditions(mSrchCond, mSrchSort, ref bookmark, out mSrchStatus);
                        if (results != null)
                            mAllSrchResults.AddRange(results);
                        else break;
                    }

                    if (mAllSrchResults.Count == 1)
                    {
                        SelectionTypeId mSelectionType = Autodesk.Connectivity.Explorer.Extensibility.SelectionTypeId.ChangeOrder;
                        LocationContext mLocation = new Autodesk.Connectivity.Explorer.Extensibility.LocationContext(mSelectionType, mAllSrchResults[0].Num);
                        e.Context.GoToLocation = mLocation;
                    }
                    else MessageBox.Show("No linked ECO found!" + Environment.NewLine + "Linked ECOs are supposed to have the property 'Initated by ECR' and according value", "Change Mangement Extension", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

            }
            catch (Exception ex)
            {
                // If something goes wrong, we don't want the exception to bubble up to Vault Explorer.
                throw new Exception("Unhandled Error: " + ex.Message);
            }
        }

        public SrchCond mCreateSrchCond(Connection connection, string mPropDispName, string mSrchText, string mAndOr, string mEntCls)
        {
            SearchRuleType mSrchRuleType = new SearchRuleType();
            if (mAndOr == "AND")
            {
                mSrchRuleType = SearchRuleType.Must;
            }
            else mSrchRuleType = SearchRuleType.May;

            PropertyService mPropSrvc = connection.WebServiceManager.PropertyService;
            PropDef mPropDefs = mPropSrvc.GetPropertyDefinitionsByEntityClassId(mEntCls).FirstOrDefault(n => n.DispName == mPropDispName);
            SrchCond mSrchCond = new SrchCond()
            {
                PropDefId = mPropDefs.Id,
                PropTyp = PropertySearchType.SingleProperty,
                SrchOper = 3, // is equal
                SrchRule = mSrchRuleType,
                SrchTxt = mSrchText
            };
            return mSrchCond;
        }

        public Autodesk.Connectivity.WebServices.ChangeOrder mCreateECObyCustEnt(Connection connection, ChangeOrderService coSvc, string coTitle, string coDescr, long[] coFileMasterIds, long[] coItemMasterIds, PropInst[] coUDPs, Email[] coEmails)
        {
            // get the default routing 
            Workflow workflow = coSvc.GetDefaultWorkflow();
            Routing[] routings = coSvc.GetRoutingsByWorkflowId(workflow.Id);
            Routing defaultRouting = routings.FirstOrDefault(n => n.IsDflt);
            if (defaultRouting == null)
                throw new Exception("No default routing");

            // get the next number in the default sequence 
            ChangeOrderNumSchm[] numberingSchemes =
                coSvc.GetNumberingSchemesByType(NumSchmType.ApplicationDefault);
            ChangeOrderNumSchm defaultNumberingScheme = numberingSchemes.FirstOrDefault();
            if (defaultNumberingScheme == null)
                throw new Exception("No default numbering scheme");
            string coNumber = coSvc.GetChangeOrderNumberBySchemeId(defaultNumberingScheme.Id);

            //update Email Subject and Body with coNumber 
            foreach (var item in coEmails)
            {
                item.Subject = "New ECO " + coNumber + " created";
                item.Body = "The ECO " + coNumber + " is available now." + Environment.NewLine + "You are supposed to complete and respond to it."
                    + Environment.NewLine + Environment.NewLine + "Sincerely - VaultExample Administrator"
                    ;
            }

            // create the new change order 
            Autodesk.Connectivity.WebServices.ChangeOrder mEcoNew = coSvc.AddChangeOrder(defaultRouting.Id, coNumber, coTitle, coDescr, DateTime.Now.AddMonths(1), coItemMasterIds, null, coFileMasterIds, coUDPs, null, null, coEmails);
            return mEcoNew;
        }

        public Autodesk.Connectivity.WebServices.ChangeOrder mCreateECObyFile(Connection mCon, ChangeOrderService coSvc, long[] mFileMasterIds, long[] mFileIds)
        {
            // get the default routing 
            Workflow workflow = coSvc.GetDefaultWorkflow();
            Routing[] routings = coSvc.GetRoutingsByWorkflowId(workflow.Id);
            Routing defaultRouting = routings.FirstOrDefault(n => n.IsDflt);
            if (defaultRouting == null)
                throw new Exception("No default routing");

            // get the next number in the default sequence 
            ChangeOrderNumSchm[] numberingSchemes =
                coSvc.GetNumberingSchemesByType(NumSchmType.ApplicationDefault);
            ChangeOrderNumSchm defaultNumberingScheme = numberingSchemes.FirstOrDefault();
            if (defaultNumberingScheme == null)
                throw new Exception("No default numbering scheme");
            string coNumber = coSvc.GetChangeOrderNumberBySchemeId(defaultNumberingScheme.Id);

            PropertyService mPropSrvc = mCon.WebServiceManager.PropertyService;
            PropDef[] mPropDefs = mPropSrvc.GetPropertyDefinitionsByEntityClassId("FILE").Where(n => n.DispName == "Title" || n.DispName == "Subject" ||
                n.DispName == "Email Recipient").ToArray();

            Dictionary<string, long> mProperties = new Dictionary<string, long>();
            int i = 0;
            foreach (var item in mPropDefs)
            {
                mProperties.Add(mPropDefs[i].DispName, mPropDefs[i].Id);
                i += 1;
            }

            long[] mPropDefIdsArray = new long[mProperties.Count()];
            mPropDefIdsArray[0] = mProperties["Title"];
            mPropDefIdsArray[1] = mProperties["Subject"];
            mPropDefIdsArray[2] = mProperties["Email Recipient"];

            PropInst[] mFileProps = mPropSrvc.GetProperties("FILE", mFileIds, mPropDefIdsArray);
            //the properties are filtered to Title and Subject, but not sorted; we need to grab the right ones...
            string mEcoTitle = "";
            var mSinglePropInst = mFileProps.Where(n => n.PropDefId == mProperties["Title"]);
            var mPropInstList = mSinglePropInst.ToList();
            if (mPropInstList[0].Val != null) mEcoTitle = mPropInstList[0].Val.ToString();

            string mEcoDescr = "";
            mSinglePropInst = mFileProps.Where(n => n.PropDefId == mProperties["Subject"]);
            mPropInstList = mSinglePropInst.ToList();
            if (mPropInstList[0].Val != null) mEcoDescr = mPropInstList[0].Val.ToString();

            string mEmailAdr = "";
            mSinglePropInst = mFileProps.Where(n => n.PropDefId == mProperties["Email Recipient"]);
            mPropInstList = mSinglePropInst.ToList();
            if (mPropInstList[0].Val != null) mEmailAdr = mPropInstList[0].Val.ToString();

            PropDef[] mEcoUdpDefs = mPropSrvc.GetPropertyDefinitionsByEntityClassId("ChangeOrder").Where(n => n.DispName == "Initiated by ECR").ToArray();
            Dictionary<string, long> mEcoProperties = new Dictionary<string, long>();
            i = 0;
            foreach (var item in mEcoUdpDefs)
            {
                mEcoProperties.Add(mEcoUdpDefs[i].DispName, mEcoUdpDefs[i].Id);
                i += 1;
            }

            Dictionary<long, FileIteration> mFiles = (Dictionary<long, FileIteration>)mCon.FileManager.GetFilesByIterationIds(mFileIds);

            long[] mEcoPropDefIdsArray = new long[mEcoProperties.Count()];
            mEcoPropDefIdsArray[0] = mEcoProperties["Initiated by ECR"];
            PropInst[] mEcoUDPs = mPropSrvc.GetProperties("ChangeOrder", new long[] { mFileIds[0] }, mEcoPropDefIdsArray);

            //alternatively use UDPs unique number Goto //alternative 2:
            #region alternative1
            //mSinglePropInst = mCustEntProps.Where(n => n.PropDefId == mProperties["Ident Number"]);
            //mPropInstList = mSinglePropInst.ToList();
            //if (mPropInstList[0].Val != null) mEcoUDPs[0].Val = mPropInstList[0].Val.ToString();
            #endregion
            #region alternative2
            DocumentService mDocSrvc = mCon.WebServiceManager.DocumentService;
            File mFile = mDocSrvc.GetFileById(mFileIds[0]);
            object mFileName = mFiles.FirstOrDefault(n => n.Key == mFile.Id);
            mEcoUDPs[0].Val = mFile.VerName;
            //object mtest = mFileName;
            #endregion

            Email[] mAllEmails = new Email[1];
            if (mEmailAdr != "")
            {
                Email mEmail = new Email();
                mEmail.ToArray = new string[] { mEmailAdr };
                mEmail.Subject = "New ECO " + coNumber + " created";
                mEmail.Body = "The ECO " + coNumber + "is available now; you are supposed to complete and open it";
                mAllEmails[0] = mEmail;
            }
            else mAllEmails = null;

            // create the new change order
            try
            {
                Autodesk.Connectivity.WebServices.ChangeOrder newCo = coSvc.AddChangeOrder(defaultRouting.Id, coNumber, mEcoTitle, mEcoDescr, DateTime.Now.AddMonths(1), null, null, mFileMasterIds, mEcoUDPs, null, null, mAllEmails);
                return newCo;
            }
            catch (Exception ex)
            {

                throw new Exception(ex.Message);
            }

        }

        public IEnumerable<CustomEntityHandler> CustomEntityHandlers()
        {
            throw new NotImplementedException();
        }

        public IEnumerable<DetailPaneTab> DetailTabs()
        {
            throw new NotImplementedException();
        }

        public IEnumerable<string> HiddenCommands()
        {
            throw new NotImplementedException();
        }

        public void OnLogOff(IApplication application)
        {
            throw new NotImplementedException();
        }

        public void OnLogOn(IApplication application)
        {
            throw new NotImplementedException();
        }

        public void OnShutdown(IApplication application)
        {
            throw new NotImplementedException();
        }

        public void OnStartup(IApplication application)
        {
            throw new NotImplementedException();
        }
    }

}
