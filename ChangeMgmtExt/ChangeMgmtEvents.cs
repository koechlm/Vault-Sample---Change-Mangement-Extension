using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.Connectivity.Extensibility.Framework;
using Autodesk.Connectivity.Explorer.Extensibility;
using Autodesk.Connectivity.WebServices;
using Autodesk.Connectivity.WebServicesTools;
using VDF = Autodesk.DataManagement.Client.Framework;
using Vault = Autodesk.DataManagement.Client.Framework.Vault;
using Autodesk.DataManagement.Client.Framework.Vault.Currency.Connections;
using Autodesk.DataManagement.Client.Framework.Vault.Currency.Entities;


namespace ChangeMgmtExt
{
    public class mChangeMgmtEvents : IWebServiceExtension
    {
        #region IWebServiceExtension Members
        public void OnLoad()
        {
            // in this case, we want to register for the GetRestrictions event for all operations

            // File Events
            DocumentService.AddFileEvents.GetRestrictions += new EventHandler<AddFileCommandEventArgs>(AddFileEvents_GetRestrictions);
            DocumentService.CheckinFileEvents.GetRestrictions += new EventHandler<CheckinFileCommandEventArgs>(CheckinFileEvents_GetRestrictions);
            DocumentService.CheckoutFileEvents.GetRestrictions += new EventHandler<CheckoutFileCommandEventArgs>(CheckoutFileEvents_GetRestrictions);
            DocumentService.DeleteFileEvents.GetRestrictions += new EventHandler<DeleteFileCommandEventArgs>(DeleteFileEvents_GetRestrictions);
            DocumentService.DownloadFileEvents.GetRestrictions += new EventHandler<DownloadFileCommandEventArgs>(DownloadFileEvents_GetRestrictions);
            DocumentServiceExtensions.UpdateFileLifecycleStateEvents.GetRestrictions += new EventHandler<UpdateFileLifeCycleStateCommandEventArgs>(UpdateFileLifecycleStateEvents_GetRestrictions);
            DocumentServiceExtensions.UpdateFileLifecycleStateEvents.Post += new EventHandler<UpdateFileLifeCycleStateCommandEventArgs>(UpdateFileLifecycleStateEvents_Post);

            // Folder Events
            DocumentService.AddFolderEvents.GetRestrictions += new EventHandler<AddFolderCommandEventArgs>(AddFolderEvents_GetRestrictions);
            DocumentService.DeleteFolderEvents.GetRestrictions += new EventHandler<DeleteFolderCommandEventArgs>(DeleteFolderEvents_GetRestrictions);

            // Item Events
            ItemService.AddItemEvents.GetRestrictions += new EventHandler<AddItemCommandEventArgs>(AddItemEvents_GetRestrictions);
            ItemService.CommitItemEvents.GetRestrictions += new EventHandler<CommitItemCommandEventArgs>(CommitItemEvents_GetRestrictions);
            ItemService.ItemRollbackLifeCycleStatesEvents.GetRestrictions += new EventHandler<ItemRollbackLifeCycleStateCommandEventArgs>(ItemRollbackLifeCycleStatesEvents_GetRestrictions);
            ItemService.DeleteItemEvents.GetRestrictions += new EventHandler<DeleteItemCommandEventArgs>(DeleteItemEvents_GetRestrictions);
            ItemService.EditItemEvents.GetRestrictions += new EventHandler<EditItemCommandEventArgs>(EditItemEvents_GetRestrictions);
            ItemService.PromoteItemEvents.GetRestrictions += new EventHandler<PromoteItemCommandEventArgs>(PromoteItemEvents_GetRestrictions);
            ItemService.UpdateItemLifecycleStateEvents.GetRestrictions += new EventHandler<UpdateItemLifeCycleStateCommandEventArgs>(UpdateItemLifecycleStateEvents_GetRestrictions);

            // Change Order Events
            ChangeOrderService.AddChangeOrderEvents.GetRestrictions += new EventHandler<AddChangeOrderCommandEventArgs>(AddChangeOrderEvents_GetRestrictions);
            ChangeOrderService.CommitChangeOrderEvents.GetRestrictions += new EventHandler<CommitChangeOrderCommandEventArgs>(CommitChangeOrderEvents_GetRestrictions);
            ChangeOrderService.DeleteChangeOrderEvents.GetRestrictions += new EventHandler<DeleteChangeOrderCommandEventArgs>(DeleteChangeOrderEvents_GetRestrictions);
            ChangeOrderService.EditChangeOrderEvents.GetRestrictions += new EventHandler<EditChangeOrderCommandEventArgs>(EditChangeOrderEvents_GetRestrictions);
            ChangeOrderService.UpdateChangeOrderLifecycleStateEvents.GetRestrictions += new EventHandler<UpdateChangeOrderLifeCycleStateCommandEventArgs>(UpdateChangeOrderLifecycleStateEvents_GetRestrictions);

            // Custom Entity Events
            CustomEntityService.UpdateCustomEntityLifecycleStateEvents.GetRestrictions += new EventHandler<UpdateCustomEntityLifeCycleStateCommandEventArgs>(UpdateCustomEntityLifecycleStateEvents_GetRestrictions);

        }

        private void UpdateCustomEntityLifecycleStateEvents_GetRestrictions(object sender, UpdateCustomEntityLifeCycleStateCommandEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void UpdateChangeOrderLifecycleStateEvents_GetRestrictions(object sender, UpdateChangeOrderLifeCycleStateCommandEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void EditChangeOrderEvents_GetRestrictions(object sender, EditChangeOrderCommandEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void DeleteChangeOrderEvents_GetRestrictions(object sender, DeleteChangeOrderCommandEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void CommitChangeOrderEvents_GetRestrictions(object sender, CommitChangeOrderCommandEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void UpdateItemLifecycleStateEvents_GetRestrictions(object sender, UpdateItemLifeCycleStateCommandEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void PromoteItemEvents_GetRestrictions(object sender, PromoteItemCommandEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void EditItemEvents_GetRestrictions(object sender, EditItemCommandEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void DeleteItemEvents_GetRestrictions(object sender, DeleteItemCommandEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void ItemRollbackLifeCycleStatesEvents_GetRestrictions(object sender, ItemRollbackLifeCycleStateCommandEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void CommitItemEvents_GetRestrictions(object sender, CommitItemCommandEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void AddItemEvents_GetRestrictions(object sender, AddItemCommandEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void DeleteFolderEvents_GetRestrictions(object sender, DeleteFolderCommandEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void AddChangeOrderEvents_GetRestrictions(object sender, AddChangeOrderCommandEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void AddFolderEvents_GetRestrictions(object sender, AddFolderCommandEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void UpdateFileLifecycleStateEvents_Post(object sender, UpdateFileLifeCycleStateCommandEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void UpdateFileLifecycleStateEvents_GetRestrictions(object sender, UpdateFileLifeCycleStateCommandEventArgs e)
        {
            ChangeMgmtExt.mClsChangeMgmtExt mStart = new mClsChangeMgmtExt();

            //mStart.mTestCommand(sender, e);
        }

        private void DownloadFileEvents_GetRestrictions(object sender, DownloadFileCommandEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void DeleteFileEvents_GetRestrictions(object sender, DeleteFileCommandEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void CheckoutFileEvents_GetRestrictions(object sender, CheckoutFileCommandEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void CheckinFileEvents_GetRestrictions(object sender, CheckinFileCommandEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void AddFileEvents_GetRestrictions(object sender, AddFileCommandEventArgs e)
        {
            throw new NotImplementedException();
        }
        #endregion
    }
}
