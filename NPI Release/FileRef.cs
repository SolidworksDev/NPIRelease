//FileRef.cs
using System;
using System.Collections;
using System.Diagnostics;
using System.Collections.Generic;
using EPDM.Interop.epdm;

namespace NPI_Release
{

    public class EDMFile
    {
        private IEdmFile5 mFile;
        private IEdmFolder5 mParentFolder;
        private IEdmState5 mCurrentState;
        private IEdmTransition10 mTransition;
        private int mLevel = 0;
        private IEdmReference10 mRefFile;
        private string mFilePath;
        private bool mUpdateDataCardFailed = false;
        private bool mCheckinFailed = false;
        private bool mChangeStateFailed = false;

        public EDMFile()
        {

        }
       
        public IEdmFile5 File
        {
            get { return mFile; }
            set { mFile = value; }
        }

        public int Level
        {
            get { return mLevel; }
            set { mLevel = value; }
        }

        public string FilePath
        {
            get { return mFilePath; }
            set { mFilePath = value; }
        }

        public IEdmFolder5 ParentFolder
        {
            get { return mParentFolder; }
            set { mParentFolder = value; }
        }

        public IEdmState5 CurrentState
        {
            get { return mCurrentState; }
            set { mCurrentState = value; }
        }

        public IEdmTransition10 Transition
        {
            get { return mTransition; }
            set { mTransition = value; }
        }

        public IEdmReference10 RefFile
        {
            get { return mRefFile; }
            set { mRefFile = value; }
        }

        public bool UpdatedDataCardFailed
        {
            get { return mUpdateDataCardFailed; }
            set { mUpdateDataCardFailed = value; }
        }

        public bool CheckinFailed
        {
            get { return mCheckinFailed; }
            set { mCheckinFailed = value; }
        }

        public bool ChangeStateFailed
        {
            get { return mChangeStateFailed; }
            set { mChangeStateFailed = value; }
        }
    }

    public class EDMFileCheckedOut
    {
        private string mFileName;
        private string mCheckedOutBy;
        private string mCheckedOutOn;

        public string FileName
        {
            get { return mFileName; }
            set { mFileName = value; }
        }

        public string CheckedOutBy
        {
            get { return mCheckedOutBy; }
            set { mCheckedOutBy = value; }
        }

        public string CheckedOutOn
        {
            get { return mCheckedOutOn; }
            set { mCheckedOutOn = value; }
        }        
    }
}