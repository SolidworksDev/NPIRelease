using System.IO;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;
using EPDM.Interop.epdm;
using EPDM.Interop.EPDMResultCode;
using Microsoft.VisualBasic;
using System.Text.RegularExpressions;
using System.Collections.Specialized;
using System.Diagnostics;
using System.Linq;
using Microsoft.VisualBasic.CompilerServices;

namespace NPI_Release
{
    public partial class NPI_Release : Form
    {
        public NPI_Release()
        {
            InitializeComponent();

            args = Environment.GetCommandLineArgs();
            
            if (args.Length > 1)
            {
                CommandLineFileName = args[1];
                tbBatchRef.Text = Path.GetFileName(CommandLineFileName);

                sw.Start();
                vault5 = new EdmVault5();
                vault7 = (IEdmVault7)vault5;
                vault8 = (IEdmVault8)vault5;
                EdmViewInfo[] Views = null;

                vault8.GetVaultViews(out Views, false);
                VaultsComboBox.Items.Clear();
                foreach (EdmViewInfo View in Views)
                {
                    VaultsComboBox.Items.Add(View.mbsVaultName);
                }
                if (VaultsComboBox.Items.Count > 0)
                {
                    VaultsComboBox.Text = (string)VaultsComboBox.Items[0];
                }
                btnAccept.Enabled = false;

                if (vault5 == null)
                {
                    vault5 = new EdmVault5();
                    vault7 = (IEdmVault7)vault5;
                }
                if (!vault5.IsLoggedIn)
                {
                    //Log into selected vault as the current user
                    vault5.LoginAuto(VaultsComboBox.Text, this.Handle.ToInt32());
                }

                validUser = CheckUser();

                if (validUser == false)
                {
                    SelectedFilesDictionary.Clear();
                    SelectedFilesList.Clear();
                    DrawingFilesList.Clear();
                    FileInfoList.Clear();
                    this.Close();
                }

                btnAccept.Enabled = false;
                btnCancel.Enabled = false;
                toolStripProgressBar1.Maximum = 20;
                getRefsWorker.RunWorkerAsync();
            }
        }

        private IEdmVault5 vault5 = null;
        private IEdmVault7 vault7 = null;
        private IEdmVault8 vault8 = null;
        private string CommandLineFileName = null;
        private bool validUser = true;

        private Dictionary<string, string> SelectedFilesDictionary = new Dictionary<string, string>();
        private List<IEdmFile5> RefFilesList = new List<IEdmFile5>();
        private List<EDMFile> SelectedFilesList = new List<EDMFile>();
        private List<EDMFile> DrawingFilesList = new List<EDMFile>();
        private List<EDMFile> FileInfoList = new List<EDMFile>();
        private List<EDMFileCheckedOut> FilesCheckedOut = new List<EDMFileCheckedOut>();
        private Stopwatch sw = new Stopwatch();
        private string ErrorsMessage = "The Following files had errors: \n";
        private string[] args = Environment.GetCommandLineArgs();

        private void NPI_Release_Load(object sender, EventArgs e)
        {
            try
            {                
                sw.Start();
                vault5 = new EdmVault5();
                vault7 = (IEdmVault7)vault5;
                vault8 = (IEdmVault8)vault5;
                EdmViewInfo[] Views = null;

                vault8.GetVaultViews(out Views, false);

                VaultsComboBox.Items.Clear();
                foreach (EdmViewInfo View in Views)
                {
                    VaultsComboBox.Items.Add(View.mbsVaultName);
                }
                if (VaultsComboBox.Items.Count > 0)
                {
                    VaultsComboBox.Text = (string)VaultsComboBox.Items[0];
                }
                btnAccept.Enabled = false;

                if (vault5 == null)
                {
                    vault5 = new EdmVault5();
                    vault7 = (IEdmVault7)vault5;
                }
                if (!vault5.IsLoggedIn)
                {
                    //Log into selected vault as the current user
                    vault5.LoginAuto(VaultsComboBox.Text, this.Handle.ToInt32());
                }

                validUser = CheckUser();

                if (validUser == false)
                {
                    SelectedFilesDictionary.Clear();
                    SelectedFilesList.Clear();
                    DrawingFilesList.Clear();
                    FileInfoList.Clear();
                    this.Close();
                }
            }

            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + " " + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + GetStackTrace(ex));
            }
        }

        private void BrowseButton_Click(object sender, EventArgs e)
        {
            OpenFile();
        }

        private void OKButton_Click(object sender, EventArgs e)
        {
            try
            {
                //Get the file paths of all of the referenced files
                //and store them in RefFilesDictionary as keys;
                //the levels where they are found in the file hierarchy 
                //are stored as values                                

                GetCheckedFiles(filesListView);

                if (SelectedFilesDictionary.Count == 0)
                {
                    MessageBox.Show("You Must Select at Lease One File", "Please Select Files", MessageBoxButtons.OK, MessageBoxIcon.Hand, MessageBoxDefaultButton.Button1);
                    return;
                }

                foreach (EDMFile f in FileInfoList)
                {
                    foreach (KeyValuePair<string, string> kvp in SelectedFilesDictionary)
                    {
                        if (Path.GetFileName(kvp.Key) == f.File.Name)
                        {
                            SelectedFilesList.Add(f);
                        }
                    }
                }

                //Because selecting a file in the Open File dialog 
                //adds the file and its references to the local cache, 
                //clear the local cache to demonstrate that the 
                //IEdmBatchListing methods don't add the files to 
                //the local cache 
                //Declare and create the IEdmClearLocalCache3 object
                IEdmClearLocalCache3 ClearLocalCache = default(IEdmClearLocalCache3);
                ClearLocalCache = (IEdmClearLocalCache3)vault7.CreateUtility(EdmUtility.EdmUtil_ClearLocalCache);
                ClearLocalCache.IgnoreToolboxFiles = true;
                ClearLocalCache.UseAutoClearCacheOption = true;

                //Declare and create the IEdmBatchListing object
                IEdmBatchListing BatchListing = default(IEdmBatchListing);
                BatchListing = (IEdmBatchListing)vault7.CreateUtility(EdmUtility.EdmUtil_BatchList);

                //Add all of the reference file paths to the 
                //IEdmClearLocalCache object to clear from the 
                //local cache and to the IEdmBatchListing object
                //to get the file information in batch mode
                foreach (KeyValuePair<string, string> KvPair in SelectedFilesDictionary)
                {
                    ClearLocalCache.AddFileByPath(KvPair.Key);
                    ((IEdmBatchListing2)BatchListing).AddFileCfg(KvPair.Key, DateTime.Now, (Convert.ToInt32(KvPair.Value)), "@", Convert.ToInt32(EdmListFileFlags.EdmList_Nothing));
                }
                //Clear the local cache of the reference files
                ClearLocalCache.CommitClear();

                statusStrip1.Text = "Updating Data Card Information...";
                CheckoutFilesWorker.RunWorkerAsync();
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + " " + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + GetStackTrace(ex));
            }
        }

        private int GetReferencedFiles(IEdmReference10 Reference, string FilePath, int Level, string ProjectName)
        {
            int count = 0;
            
            try

            {
                bool Top = false;
                if (Reference == null)
                {
                    Top = true;
                    IEdmFile5 File = null;
                    IEdmFolder5 ParentFolder = null;
                    File = vault5.GetFileFromPath(FilePath, out ParentFolder);
                    if (File != null)
                    {

                        EDMFile anEDMFile = new EDMFile();
                        anEDMFile.File = File;
                        anEDMFile.Level = Level;
                        anEDMFile.FilePath = FilePath;
                        anEDMFile.ParentFolder = ParentFolder;
                        FileInfoList.Add(anEDMFile);
                        Reference = (IEdmReference10)File.GetReferenceTree(ParentFolder.ID);
                        anEDMFile.RefFile = Reference;
                        count += GetReferencedFiles(Reference, "", Level + 1, ProjectName);
                    }
                    else
                    {
                        MessageBox.Show("Unable to open file: " + Path.GetFileName(FilePath), "File Not Found!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return 0;
                    }
                }
                else
                {
                    //Execute this code when this function is called recursively; 
                    //i.e., this is not the top-level IEdmReference in the tree

                    //Recursively traverse the references
                    IEdmPos5 pos = default(IEdmPos5);

                    IEdmReference10 Reference2 = (IEdmReference10)Reference;
                    pos = Reference2.GetFirstChildPosition4(ProjectName, Top, true, false, (int)EdmRefFlags.EdmRef_File, "", 0);

                    IEdmReference10 aFile = default(IEdmReference10);

                    while ((!pos.IsNull))
                    {

                        aFile = (IEdmReference10)Reference.GetNextChild(pos);
                        IEdmPos5 ppos = aFile.File.GetFirstFolderPosition();
                        IEdmFolder5 ParentFolder = default(IEdmFolder5);

                        if (!ppos.IsNull)
                        {
                            ParentFolder = aFile.File.GetNextFolder(ppos);
                        }

                        EDMFile anEDMFile = new EDMFile();
                        anEDMFile.File = aFile.File;
                        anEDMFile.RefFile = aFile;
                        anEDMFile.Level = Level;
                        anEDMFile.FilePath = aFile.FoundPath;
                        anEDMFile.ParentFolder = ParentFolder;
                        FileInfoList.Add(anEDMFile);
                        count++;
                        string fileName = aFile.Name;
                        if (IsComponent(fileName))
                        {
                            getRefsWorker.ReportProgress(count, fileName);
                        }
                        count += GetReferencedFiles(aFile, "", Level + 1, ProjectName);
                        }
                    }
                }

            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + " " + ex.Message);
            }
            
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + GetStackTrace(ex));
            }
            return count;
        }

        private void GetReferencingFiles(EDMFile File)
        {
            try
            {
                IEdmPos5 ppos = File.File.GetFirstFolderPosition();
                IEdmFolder5 ParentFolder = File.File.GetNextFolder(ppos);

                IEdmReference10 Reference = (IEdmReference10)File.File.GetReferenceTree(ParentFolder.ID);
                IEdmReference10 Reference2 = (IEdmReference10)Reference;
                IEdmPos5 pos = default(IEdmPos5);
                IEdmReference10 drawing = default(IEdmReference10);
                pos = Reference2.GetFirstParentPosition2(0, false, (int)EdmRefFlags.EdmRef_File);

                while ((!pos.IsNull))
                {
                    drawing = (IEdmReference10)Reference.GetNextParent(pos);

                    if (DrawingFilesList.Count == 0)
                    {
                        EDMFile anEDMFile = new EDMFile();
                        anEDMFile.File = drawing.File;
                        anEDMFile.Level = File.Level;
                        anEDMFile.FilePath = drawing.FoundPath;
                        anEDMFile.ParentFolder = ParentFolder;
                        DrawingFilesList.Add(anEDMFile);
                        return;
                    }
                    else
                    {
                        bool found = DrawingFilesList.Any(draw => draw.FilePath == drawing.FoundPath);

                        if (!found)
                        {
                            string extension = Path.GetExtension(drawing.Name).ToLower();
                            if (extension == ".slddrw")
                            {
                                EDMFile anEDMFile = new EDMFile();
                                anEDMFile.File = drawing.File;
                                anEDMFile.Level = File.Level;
                                anEDMFile.FilePath = drawing.FoundPath;
                                anEDMFile.ParentFolder = ParentFolder;
                                DrawingFilesList.Add(anEDMFile);
                                return;
                            }
                        }
                    }
                }
            }

            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + " " + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + GetStackTrace(ex));
            }
        }

        private void GetDrawingFiles(List<EDMFile> SelectedFiles)
        {
            try
            {
                int progress = 1;
                CheckoutFilesWorker.ReportProgress(-1, SelectedFiles.Count);

                foreach (EDMFile sf in SelectedFiles)
                {
                    GetReferencingFiles(sf);
                    CheckoutFilesWorker.ReportProgress(progress);
                    ++progress;
                }
            }

            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + " " + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + GetStackTrace(ex));
            }
        }

        private void BatchCheckout(List<EDMFile> SelectedFiles)
        {
            EdmSelItem[] ppoSelection = new EdmSelItem[SelectedFiles.Count];
            IEdmBatchGet batchGetter = (IEdmBatchGet)vault7.CreateUtility(EdmUtility.EdmUtil_BatchGet);
            IEdmFolder5 aFolder;
            IEdmPos5 aPos;
            int i = 0;

            CheckoutFilesWorker.ReportProgress(-1, SelectedFiles.Count);

            try
            {
                foreach (EDMFile sf in SelectedFiles)
                {
                    aPos = sf.File.GetFirstFolderPosition();
                    aFolder = sf.File.GetNextFolder(aPos);
                    ppoSelection[i] = new EdmSelItem();
                    ppoSelection[i].mlDocID = sf.File.ID;
                    ppoSelection[i].mlProjID = aFolder.ID;
                    i = i + 1;
                    CheckoutFilesWorker.ReportProgress(i);
                }
                batchGetter.AddSelection((EdmVault5)vault5, ref ppoSelection);
                batchGetter.CreateTree(0, (int)EdmGetCmdFlags.Egcf_Lock);
                batchGetter.GetFiles(0, null);
                CheckoutFilesWorker.ReportProgress(0);
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + " " + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + GetStackTrace(ex));
            }
        }

        private void UpdateDataCard(List<EDMFile> SelectedFiles)
        {
            int progress = 0;
            EDMFile FileException = default(EDMFile);
            CheckoutFilesWorker.ReportProgress(-1, SelectedFiles.Count);

            try
            {
                IEdmBatchUpdate2 Update = default(IEdmBatchUpdate2);
                Update = (IEdmBatchUpdate2)vault7.CreateUtility(EdmUtility.EdmUtil_BatchUpdate);

                foreach (EDMFile sf in SelectedFiles)
                {
                    FileException = sf;
                    IEdmVariableMgr5 varMgr = default(IEdmVariableMgr5);
                    varMgr = (IEdmVariableMgr5)sf.File.Vault;

                    int ECONumberID = 0;
                    int RevisionDescriptionID = 0;

                    ECONumberID = varMgr.GetVariable("ECO Number").ID;
                    RevisionDescriptionID = varMgr.GetVariable("RevisionDescription").ID;

                    if (RevisionDescriptionID != 0)
                    {
                        Update.SetVar(sf.File.ID, RevisionDescriptionID, tbRevisionDescription.Text, "", (int)EdmBatchFlags.EdmBatch_Nothing);
                    }

                    if (ECONumberID != 0)
                    {
                        Update.SetVar(sf.File.ID, ECONumberID, tbECONumber.Text, "", (int)EdmBatchFlags.EdmBatch_Nothing);
                    }

                    progress++;
                    CheckoutFilesWorker.ReportProgress(progress);
                }

                //Write all the card variable values to the database
                EdmBatchError2[] Errors = null;
                int errorSize = 0;
                errorSize = Update.CommitUpdate(out Errors, null);
            }

            catch (System.Runtime.InteropServices.COMException ex)
            {
                //MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + " " + ex.Message);
                FileException.UpdatedDataCardFailed = true;
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message + "\n" + GetStackTrace(ex));
                FileException.UpdatedDataCardFailed = true;
            }
        }

        private void BatchCheckin(List<EDMFile> SelectedFiles)
        {
            int i = 0;
            EDMFile FileException = default(EDMFile);
            IEdmFolder5 aFolder;
            IEdmPos5 aPos;
            EdmSelItem[] ppoSelection = new EdmSelItem[SelectedFiles.Count];
            IEdmBatchUnlock batchUnlock = (IEdmBatchUnlock)vault7.CreateUtility(EdmUtility.EdmUtil_BatchUnlock);

            CheckinFilesWorker.ReportProgress(-1, SelectedFiles.Count);

            try
            {
                foreach (EDMFile sf in SelectedFiles)
                {
                    FileException = sf;
                    aPos = sf.File.GetFirstFolderPosition();
                    aFolder = sf.File.GetNextFolder(aPos);
                    ppoSelection[i] = new EdmSelItem();
                    ppoSelection[i].mlDocID = sf.File.ID;
                    ppoSelection[i].mlProjID = aFolder.ID;
                    i++;
                    CheckinFilesWorker.ReportProgress(i);
                }

                batchUnlock.AddSelection((EdmVault5)vault5, ref ppoSelection);
                batchUnlock.CreateTree(0, (int)EdmUnlockBuildTreeFlags.Eubtf_ShowCloseAfterCheckinOption + (int)EdmUnlockBuildTreeFlags.Eubtf_MayUnlock);
                batchUnlock.Comment = "Release to Production ECO";
                batchUnlock.UnlockFiles(0, null);
                CheckinFilesWorker.ReportProgress(0);
            }

            catch (System.Runtime.InteropServices.COMException ex)
            {
                //MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + " " + ex.Message);
                FileException.CheckinFailed = true;
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message + "\n" + GetStackTrace(ex));
                FileException.CheckinFailed = true;
            }
        }

        private void DisplayState(List<EDMFile> SelectedFiles)
        {
            foreach (EDMFile sf in SelectedFiles)
            {
                IEdmFile10 aFile = (IEdmFile10)sf.File;
                string CurrentStateName = aFile.CurrentState.Name;
                IEdmState5 workflowState = aFile.CurrentState;
                IEdmPos5 pos = default(IEdmPos5);
                pos = workflowState.GetFirstTransitionPosition();

                if (!pos.IsNull)
                {
                    MessageBox.Show("No next state");
                }

                IEdmTransition10 transition = default(IEdmTransition10);
                transition = (IEdmTransition10)workflowState.GetNextTransition(pos);

                string message = "Current state for " + aFile.Name + ": " + "\n";
                message = message + " * " + workflowState.Name + "\n" + "\n";
                message = message + "Next possable state transitions for " + aFile.Name + ":" + "\n";
                message = message + " * " + transition.Name + "\n";

                while ((!pos.IsNull))
                {
                    transition = (IEdmTransition10)workflowState.GetNextTransition(pos);
                    message = message + " * " + transition.Name + "\n";
                }

                MessageBox.Show(message);
            }
        }

        private void ChangeState(List<EDMFile> SelectedFiles)
        {
            EDMFile FileException = default(EDMFile);
            IEdmBatchChangeState5 batchChanger;
            batchChanger = (IEdmBatchChangeState5)vault7.CreateUtility(EdmUtility.EdmUtil_BatchChangeState);
            bool retVal;
            int progress = 1;
            CheckinFilesWorker.ReportProgress(-1, SelectedFiles.Count);

            try
            {

                foreach (EDMFile sf in SelectedFiles)
                 {
                     batchChanger.AddFile(sf.File.ID, sf.ParentFolder.ID);
                     CheckinFilesWorker.ReportProgress(progress);
                     progress++;
                 }
    
                 batchChanger.Comment = "Automated State Change";
                 if (SelectedFiles[0].CurrentState.Name == "Under Change")
                 {

                     retVal = batchChanger.CreateTree("Request Eng Change Review");
                 }
                 else
                 {
                     retVal = batchChanger.CreateTree("Request Eng Release Review");
                 }
                 batchChanger.ChangeState(0);
            }

            catch (System.Runtime.InteropServices.COMException ex)
            {
                //MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + " " + ex.Message);
                FileException.ChangeStateFailed = true;
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message + "\n" + GetStackTrace(ex));
                FileException.ChangeStateFailed = true;
            }
        }

        private void GetWorkflowState(EDMFile SelectedFile)
        {
            IEdmState5 workflowState = default(IEdmState5);
            workflowState = SelectedFile.File.CurrentState;
            IEdmPos5 pos = default(IEdmPos5);
            pos = workflowState.GetFirstTransitionPosition();
            string workflowName = "";
            int workflowID = 0;

            if (pos.IsNull)
            {
                MessageBox.Show("No next state");
                return;
            }

            while (!pos.IsNull)
            {
                IEdmTransition10 ts = (IEdmTransition10)workflowState.GetNextTransition(pos);
                workflowName = ts.Name;
                workflowID = ts.ToStateID;

                if (workflowName == "Request Eng Release Review")
                {
                    SelectedFile.Transition = ts;
                }
            }
            SelectedFile.CurrentState = workflowState;
        }       

        private void PopulateListView(ListView lv)
        {
            try
            {
                lv.Columns.Add("File Name");
                lv.SmallImageList = imageList1;
                ColumnHeader col;

                col = lv.Columns[0];
                col.Width = -2;

                ListViewItem lvi;

                lv.BeginUpdate();

                foreach (EDMFile aFile in FileInfoList)
                {

                    if (aFile.File.Name != "Cut-List-Item1")
                    {
                        if (IsInDesign(aFile.RefFile, aFile.Level))
                        {
                            if (aFile.Level == 0)
                            {
                                lvi = lv.Items.Add(aFile.File.Name);
                                lvi.ImageIndex = 0;
                                lvi.Checked = true;
                            }
                            else if (IsComponent(aFile.File.Name))
                            {
                                string ext = Path.GetExtension(aFile.File.Name);
                                lvi = lv.Items.Add(aFile.File.Name);
                                if (ext == ".SLDASM" || ext == ".sldasm")
                                {
                                    lvi.ImageIndex = 0;
                                }
                                else
                                {
                                    lvi.ImageIndex = 1;
                                }
                            }
                        }
                    }
                }
                lv.Sort();
                lv.EndUpdate();
            }

            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + " " + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + GetStackTrace(ex));
            }
        }

        private void GetCheckedFiles(ListView lv)
        {
            try
            {
                foreach (ListViewItem lvi in lv.CheckedItems)
                {
                    if (!SelectedFilesDictionary.Any(v => v.Key == lvi.Text))
                    {
                        SelectedFilesDictionary.Add(lvi.Text, "1");
                    }
                }
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + " " + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + GetStackTrace(ex));
            }
            }

        private void Cancel_Click(object sender, EventArgs e)
        {
            ExitApp();
        }

        private string GetStackTrace(Exception ex)
        {
            StackTrace st = new StackTrace(ex, true);
            StackFrame[] frames = st.GetFrames();
            string stackTrace = "";

            // Iterate over the frames extracting the information you need
            foreach (StackFrame frame in frames)
            {
                stackTrace = frame.GetFileName() + " " + frame.GetMethod().Name + " " + frame.GetFileLineNumber() + " " + frame.GetFileColumnNumber() + "\n";

                Console.WriteLine("{0}:{1}({2},{3})", frame.GetFileName(), frame.GetMethod().Name, frame.GetFileLineNumber(), frame.GetFileColumnNumber());
            }

            return stackTrace;
        }

        private bool IsInDesign(IEdmReference10 refFile, int level)
        {
            try
            {
                IEdmVault7 vault2 = null;
                vault2 = (IEdmVault7)vault5;
                //Declare and create the IEdmBatchListing object
                IEdmBatchListing BatchListing = default(IEdmBatchListing);
                BatchListing = (IEdmBatchListing)vault2.CreateUtility(EdmUtility.EdmUtil_BatchList);

                ((IEdmBatchListing2)BatchListing).AddFileCfg(refFile.FoundPath, DateTime.Now, level, "@", Convert.ToInt32(EdmListFileFlags.EdmList_Nothing));

                EdmListCol[] BatchListCols = null;
                ((IEdmBatchListing2)BatchListing).CreateListEx("\nECO Number\nRevisionDescription", Convert.ToInt32(EdmCreateListExFlags.Edmclef_GetDrawings), ref BatchListCols, null);

                //Get the generated file information
                EdmListFile[] BatchListFiles = null;
                BatchListing.GetFiles(ref BatchListFiles);

                if (BatchListFiles[0].moCurrentState.mbsStateName == "In Design")
                {
                    return true;
                }
                else if (BatchListFiles[0].moCurrentState.mbsStateName == "Under Change")
                {
                    return true;
                }
                return false;
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + " " + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + GetStackTrace(ex));
            }
            return false;
        }

        private bool IsFileLocked(IEdmFile5 aFile)
        {
            if (aFile.IsLocked)
            {
                EDMFileCheckedOut aCheckedOutFile = new EDMFileCheckedOut();

                aCheckedOutFile.FileName = aFile.Name;
                aCheckedOutFile.CheckedOutBy = aFile.LockedByUser.Name;
                aCheckedOutFile.CheckedOutOn = aFile.LockedOnComputer;
                FilesCheckedOut.Add(aCheckedOutFile);

                return true;
            }
            return false;
        }

        private bool IsComponent(string name)
        {

            if (name.Length <= 3)
            {
                return false;
            }

            string prefix = name.Substring(0, 3);

            switch (prefix)
            {
                case "100":
                    return true;
                case "101":
                    return true;
                case "105":
                    return true;
                case "110":
                    return true;
                case "115":
                    return true;
                case "125":
                    return true;
                case "154":
                    return true;
                case "200":
                    return true;
                case "205":
                    return true;
                case "225":
                    return true;
                case "240":
                    return true;
                case "250":
                    return true;
                case "251":
                    return true;
                case "275":
                    return true;
                case "333":
                    return true;
                case "334":
                    return true;
                case "343":
                    return true;
                case "344":
                    return true;
                case "345":
                    return true;
                case "346":
                    return true;
                case "347":
                    return true;
                case "370":
                    return true;
                case "400":
                    return true;
                case "405":
                    return true;
                case "425":
                    return true;
                case "426":
                    return true;
                case "444":
                    return true;
                case "445":
                    return true;
                case "500":
                    return true;
                case "502":
                    return true;
                case "505":
                    return true;
                case "506":
                    return true;
                case "507":
                    return true;
                case "508":
                    return true;
                case "509":
                    return true;
                case "510":
                    return true;
                case "511":
                    return true;
                case "512":
                    return true;
                case "513":
                    return true;
                case "514":
                    return true;
                case "515":
                    return true;
                case "600":
                    return true;
                case "605":
                    return true;
                case "615":
                    return true;
                case "602":
                    return true;
                case "625":
                    return true;
                case "630":
                    return true;
                case "634":
                    return true;
                case "640":
                    return true;
                case "650":
                    return true;
                case "655":
                    return true;
                case "665":
                    return true;
                case "670":
                    return true;
                case "695":
                    return true;
                case "696":
                    return true;
                case "697":
                    return true;
                case "698":
                    return true;
                case "699":
                    return true;
                case "700":
                    return true;
                case "705":
                    return true;
                case "710":
                    return true;
                case "717":
                    return true;
                case "718":
                    return true;
                case "740":
                    return true;
                case "741":
                    return true;
                case "743":
                    return true;
                case "744":
                    return true;
                case "750":
                    return true;
                case "755":
                    return true;
                case "777":
                    return true;
                case "790":
                    return true;
                case "800":
                    return true;
                case "801":
                    return true;
                case "802":
                    return true;
                case "803":
                    return true;
                case "804":
                    return true;
                case "805":
                    return true;
                case "806":
                    return true;
                case "807":
                    return true;
                case "808":
                    return true;
                case "809":
                    return true;
                case "810":
                    return true;
                case "816":
                    return true;
                case "817":
                    return true;
                case "820":
                    return true;
                case "821":
                    return true;
                case "822":
                    return true;
                case "823":
                    return true;
                case "824":
                    return true;
                case "835":
                    return true;
                case "826":
                    return true;
                case "827":
                    return true;
                case "828":
                    return true;
                case "829":
                    return true;
                case "840":
                    return true;
                case "850":
                    return true;
                case "860":
                    return true;
                case "870":
                    return true;
                case "871":
                    return true;
                case "880":
                    return true;
                case "888":
                    return true;
                case "900":
                    return true;
                case "909":
                    return true;
                case "910":
                    return true;
                case "913":
                    return true;
                case "915":
                    return true;
                default:
                    if (Operators.LikeString(prefix, "M-*", CompareMethod.Text))
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
            }
        }

        private bool CheckForErrors(List<EDMFile> SelectedFiles)
        {
            bool ErrorsFound = false;

            foreach (EDMFile sf in SelectedFiles)
            {
                if (sf.UpdatedDataCardFailed == true)
                {
                    ErrorsMessage += "Error updating the data card in: " + sf.File.Name + "\n";
                    ErrorsFound = true;
                }

                if (sf.CheckinFailed == true)
                {
                    ErrorsMessage += "Error checking in file: " + sf.File.Name + "\n";
                    ErrorsFound = true;
                }

                if (sf.ChangeStateFailed == true)
                {
                    ErrorsMessage += "Error during state change: " + sf.File.Name + "\n";
                    ErrorsFound = true;
                }
            }
            return ErrorsFound;
        }

        private bool CheckUser()
        {

            IEdmUserMgr7 userMgr = default(IEdmUserMgr7);
            userMgr = (IEdmUserMgr7)vault7.CreateUtility(EdmUtility.EdmUtil_UserMgr);

            string userName = System.Environment.UserName;
            List<string> usersGroups = new List<string>();

            IEdmUser8 CurrentUser = (IEdmUser8)userMgr.GetUser(userName);
            object[] groups = default(object[]);

            CurrentUser.GetGroupMemberships(out groups);
            string message = "Current users: " + userName + "\n Does not have permission to use this tool";

            foreach (object g in groups)
            {
                IEdmUserGroup7 group = (IEdmUserGroup7)g;
                usersGroups.Add(group.Name);
            }

            if (!usersGroups.Contains("engineer"))
            {
                MessageBox.Show(message);
                return false;
            }
            return true;
        }

        private void OpenFile()
        {
            try
            {

                if (vault5 == null)
                {
                    vault5 = new EdmVault5();
                    vault7 = (IEdmVault7)vault5;
                }
                if (!vault5.IsLoggedIn)
                {
                    //Log into selected vault as the current user
                    vault5.LoginAuto(VaultsComboBox.Text, this.Handle.ToInt32());
                }

                //Set the filter to only show Solidworks .sldasm files
                OpenFileDialog.Filter = "Assembly files *.sldasm | *.sldasm";
                //Set the initial directory in the Open dialog
                OpenFileDialog.InitialDirectory = vault5.RootFolderPath;

                //Show the Open dialog
                System.Windows.Forms.DialogResult DialogResult;
                DialogResult = OpenFileDialog.ShowDialog();
                //If the user didn't click Open, exit
                if (!(DialogResult == System.Windows.Forms.DialogResult.OK))
                {
                    return;
                }

                tbBatchRef.Text = Path.GetFileName(OpenFileDialog.FileName);

                btnAccept.Enabled = false;
                btnCancel.Enabled = false;
                toolStripProgressBar1.Maximum = 20;
                getRefsWorker.RunWorkerAsync();
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + " " + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + GetStackTrace(ex));
            }
        }

        private void ExitApp()
        {
            SelectedFilesDictionary.Clear();
            SelectedFilesList.Clear();
            DrawingFilesList.Clear();
            FileInfoList.Clear();
            this.Close();
        }

        #region background_workers

        private void getRefsWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                toolStripStatusLabel1.Text = "Building File List...";
                if (CommandLineFileName != null)
                {
                    GetReferencedFiles(null, CommandLineFileName, 0, "A");
                    toolStripStatusLabel1.Text = "Building File List...";
                }
                else
                {
                    GetReferencedFiles(null, OpenFileDialog.FileName, 0, "A");
                    toolStripStatusLabel1.Text = "Building File List...";
                }
            }

            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + " " + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + GetStackTrace(ex));
            }
        }

        private void getRefsWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            try
            {
                if (FilesCheckedOut.Count > 0)
                {
                    string CheckedOutFilesList = "";
                    foreach (EDMFileCheckedOut aFile in FilesCheckedOut)
                    {
                        CheckedOutFilesList = aFile.FileName + " Checked out by: " + aFile.CheckedOutBy + "On Computer: " + aFile.CheckedOutOn + "\n";
                    }

                    DialogResult dialogResult = MessageBox.Show("The following files are checked out\n\n" + CheckedOutFilesList + "\nDo you want to continue?", "Warning!", MessageBoxButtons.YesNo);

                    if (dialogResult == DialogResult.Yes)
                    {
                        filesListView.Sort();
                        PopulateListView(filesListView);
                        btnAccept.Enabled = true;
                        btnCancel.Enabled = true;
                        toolStripStatusLabel1.Text = "Select File to Update.";
                        toolStripProgressBar1.Value = 0;
                    }
                    else if (dialogResult == DialogResult.No)
                    {
                        RefFilesList.Clear();
                        btnAccept.Enabled = false;
                        btnCancel.Enabled = true;
                        toolStripStatusLabel1.Text = "Operation Canceled by User.";
                        toolStripProgressBar1.Value = 0;
                        return;
                    }
                }
                else
                {                    
                    filesListView.Sort();
                    PopulateListView(filesListView);
                    btnAccept.Enabled = true;
                    btnCancel.Enabled = true;
                    toolStripStatusLabel1.Text = "Select File to Update.";
                    toolStripProgressBar1.Value = 0;
                }
            }

            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + " " + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + GetStackTrace(ex));
            }
        }

        private void getRefsWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            try
            {
                string fileName = e.UserState.ToString();

                if (fileName != null)
                {
                    toolStripStatusLabel1.Text = fileName;
                }

                if (e.ProgressPercentage == 1)
                {
                    toolStripProgressBar1.Maximum = 20;
                }
                else
                {
                    if (toolStripProgressBar1.Maximum > (int)e.ProgressPercentage)
                    {
                        toolStripProgressBar1.Value = (int)e.ProgressPercentage;
                    }
                    else
                    {
                        toolStripProgressBar1.Maximum = (int)e.ProgressPercentage * 2;
                        toolStripProgressBar1.Value = (int)e.ProgressPercentage;
                    }

                }
            }

            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + " " + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + GetStackTrace(ex));
            }
        }

        private void CheckoutFilesWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                toolStripStatusLabel1.Text = "Getting Drawing Files...";
                GetDrawingFiles(SelectedFilesList);
                foreach (EDMFile df in DrawingFilesList)
                {
                    SelectedFilesList.Add(df);
                }


                //foreach (EDMFiles sf in SelectedFilesList)
                //{
                //    GetWorkflowState(sf);
                //}

                toolStripStatusLabel1.Text = "Checking out Files...";
                BatchCheckout(SelectedFilesList);
                toolStripStatusLabel1.Text = "Updating Data Cards...";
                UpdateDataCard(SelectedFilesList);
            }

            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + " " + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + GetStackTrace(ex));
            }
        }

        private void CheckoutFilesWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            try
            {
                CheckinFilesWorker.RunWorkerAsync();
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + " " + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + GetStackTrace(ex));
            }
        }

        private void CheckoutFilesWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage == -1)
            {
                toolStripProgressBar1.Maximum = Convert.ToInt32(e.UserState);
            }
            else
            {
                toolStripProgressBar1.Value = (int)e.ProgressPercentage;
            }
        }

        private void CheckinFilesWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                toolStripStatusLabel1.Text = "Checking in Files...";
                BatchCheckin(SelectedFilesList);
                toolStripStatusLabel1.Text = "Moving State to Request Eng Release Review...";
                ChangeState(SelectedFilesList);
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + " " + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + GetStackTrace(ex));
            }
        }

        private void CheckinFilesWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            bool FoundErrors = CheckForErrors(SelectedFilesList);

            if (FoundErrors)
            {
                MessageBox.Show(ErrorsMessage);
            }

            try
            {
                sw.Stop();
                TimeSpan ts = sw.Elapsed;
                toolStripStatusLabel1.Text = SelectedFilesList.Count.ToString() + " Files Successfully Updated!    " + "Elapsed Time: " + ts.ToString(@"m\:ss");
                toolStripProgressBar1.Value = 0;
                btnAccept.Enabled = false;
                btnCancel.Enabled = true;
                tbBatchRef.Text = "";
                SelectedFilesDictionary.Clear();
                FileInfoList.Clear();
                DrawingFilesList.Clear();
                SelectedFilesList.Clear();
                filesListView.Clear();
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + " " + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + GetStackTrace(ex));
            }
        }

        private void CheckinFilesWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage == -1)
            {
                toolStripProgressBar1.Maximum = Convert.ToInt32(e.UserState);
            }
            else
            {
                toolStripProgressBar1.Value = (int)e.ProgressPercentage;
            }
        }

        #endregion

        #region testcode

        private void TraverseGroups()
        {

            //Declare and create an instance of IEdmVault5 object
            IEdmVault5 vault = new EdmVault5();
            //Log into selected vault as the current user
            vault.LoginAuto(VaultsComboBox.Text, this.Handle.ToInt32());

            //Declare an IEdmUserMgr5 object
            IEdmUserMgr5 UserMgr = default(IEdmUserMgr5);
            //The IEdmUserMgr5 interface is implemented by the
            //same class as the IEdmVault5 interface,
            //so assign the value of the IEdmVault5 object
            UserMgr = (IEdmUserMgr5)vault;

            string Groups = "";
            IEdmPos5 UserGroupPos = default(IEdmPos5);
            UserGroupPos = UserMgr.GetFirstUserGroupPosition();
            while (!UserGroupPos.IsNull)
            {
                IEdmUserGroup5 UserGroup = default(IEdmUserGroup5);
                UserGroup = UserMgr.GetNextUserGroup(UserGroupPos);
                Groups = Groups + UserGroup.Name + "\n";
            }
            MessageBox.Show(Groups, vault.Name + " Vault Groups", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        #endregion

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            NPIReleaseAboutBox aboutBox = new NPIReleaseAboutBox();
            aboutBox.Show();
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFile();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ExitApp();
        }
    }
}
