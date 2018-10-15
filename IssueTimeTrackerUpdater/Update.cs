using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace IssueTimeTrackerUpdater
{
    public class Update
    {
        public string Version { get; set; }
        public string ReleaseDate { get; set; }
        public string ReleaseNotes { get; set; }
        public bool RequiredUpdate { get; set; } //Won't allow access to the program unless updated. If false, the update is optional
        public bool VerifyUpdate { get; set; }
        public bool Skip { get; set; }
        public List<UpdateData> UpdateData = new List<UpdateData>();
        public List<DownloadFile> Files = new List<DownloadFile>();
    }
}
