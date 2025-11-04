using HashidsNet;

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

        ///// <summary>
        ///// ClickUp Space ID
        ///// for SFC Projects: https://app.clickup.com/9010105092/v/s/90110035866
        ///// defined in the ClickUp URL.
        ///// </summary>
        //public const string CLICKUP_SPACE_ID = "90110035866";    // SFC Projects: https://app.clickup.com/9010105092/v/s/90110035866

        ///// <summary>
        ///// ClickUp Folder ID
        ///// for Bug and Issue Tracking: https://app.clickup.com/9010105092/v/f/90112247343/90110035866
        ///// defined in the ClickUp URL.
        ///// </summary>
        //public const string CLICKUP_FOLDER_ID = "90112247343";   // Bug and Issue Tracking: https://app.clickup.com/9010105092/v/f/90112247343/90110035866

        /// <summary>
        /// ClickUp List ID
        /// for Epicor-Kinetic Wiki Images: https://app.clickup.com/9010105092/v/l/6-901112215280-1
        /// defined in the ClickUp URL.
        /// </summary>
        public const string CLICKUP_LIST_ID = "901112215280";

        #endregion ClickUp API Constants

        #region --------------- Methods ---------------

        // Add any global methods here if needed in the future.

        internal static string CreateUniqueImageId(string saltString)
        {
            var hashids = new Hashids(saltString, minHashLength: 6);
            string id = hashids.Encode(123);
            // Example: "j0gW4e"

            return id;
        }

        /// <summary>
        /// Convert image EMU's to inch.
        /// Where EMU = English Metric Units.
        /// </summary>
        /// <param name="emu"></param>
        /// <returns></returns>
        public static float ConvertEMUtoInch(this long emu)
        {
            return (float)(emu / 914400.0);
        }

        /// <summary>
        /// Convert image EMU's to pixels.
        /// Where EMU = English Metric Units.
        /// </summary>
        /// <param name="emu"></param>
        /// <returns></returns>
        public static float ConvertEMUtoPixels(this long emu)
        {
            return (float)(emu / 9525.0);
        }

        #endregion --------------- Methods ---------------
    }
}