using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace ClickUpDocumentImporter.Helpers
{
    internal static class Globals
    {


        #region ClickUp API Constants

        /// <summary>
        /// The ClickUp API key for authentication.
        /// </summary>
        public const string CLICKUP_API_KEY = "pk_57274748_RVNSTG0AVZLAVQYNOEHO00KB395OGFA6";

        /// <summary>
        /// ClickUp Workspace ID
        /// for Workspace: https://app.clickup.com/9010105092/v/l/8cgpjr4-36351?pr=90110035866
        /// defined in the ClickUp URL.
        /// </summary>
        public const string CLICKUP_WORKSPACE_ID = "9010105092"; // Workspace: https://app.clickup.com/9010105092/v/l/8cgpjr4-36351?pr=90110035866

        /// <summary>
        /// ClickUp Space ID
        /// for SFC Projects: https://app.clickup.com/9010105092/v/s/90110035866
        /// defined in the ClickUp URL.
        /// </summary>
        public const string CLICKUP_SPACE_ID = "90110035866";    // SFC Projects: https://app.clickup.com/9010105092/v/s/90110035866

        /// <summary>
        /// ClickUp Folder ID
        /// for Bug and Issue Tracking: https://app.clickup.com/9010105092/v/f/90112247343/90110035866
        /// defined in the ClickUp URL.
        /// </summary>
        public const string CLICKUP_FOLDER_ID = "90112247343";   // Bug and Issue Tracking: https://app.clickup.com/9010105092/v/f/90112247343/90110035866

        /// <summary>
        /// ClickUp List ID
        /// for Backlog: https://app.clickup.com/9010105092/v/li/901104177565
        /// defined in the ClickUp URL.
        /// </summary>
        public const string CLICKUP_LIST_ID = "901104177565";    // Backlog: https://app.clickup.com/9010105092/v/li/901104177565

        #endregion ClickUp API Constants


        #region  --------------- Methods ---------------
        // Add any global methods here if needed in the future.

        #endregion  --------------- Methods ---------------
    }
}
