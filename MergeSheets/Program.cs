using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Smartsheet.Api;
using Smartsheet.Api.Models;
using Smartsheet.Api.OAuth;

namespace MergeSheets
{
    class Program
    {
        // MergeSheets [source sheet id] [destination sheet id] [key column name]
        // iterate through the rows in the source sheet
        // try to match the value in the key column with a row in the destination sheet
        // if there is a match - overwrite the values in the destination sheet with the source sheet
        // if there is no match - create a new row

        private static long source_id = 0;
        private static long dest_id = 0;
        private static string key_column = "";

        private static string OAuth_client_id = "undhiisshsfdtjx6wn";
        private static string OAuth_app_secret = "1oft5chvdgstgxi4oxd";

        static void Main(string[] args)
        {

            bool fSuccess = false;

            // extract parameters
            if ( args.Length >= 3 )
            {
                if ( long.TryParse(args[0], out source_id) && long.TryParse(args[1], out dest_id) )
                {
                    key_column = args[2];
                }
            }

            if ( key_column.Length == 0 )
            {
                Console.Write ("Invalid parameters: ");
                foreach (var item in args)
                {
                    Console.Write("{0} ", item);
                }
                Console.WriteLine("");
            }
            else
            {
                Console.WriteLine("params OK");
                try
                {
                    // hardcoded params
                    source_id = 8096118883018628;
                    dest_id =   7656726548768644;
                    key_column = "ID";

                    Token token = OAuthDesktopApp();   // returns a valid token or null
                    if (token != null)
                    {
                        fSuccess = MergeSheets(token);
                    }

                }
                catch (Exception e)
                {
                    Console.WriteLine("\n*** Exception in MergeSheets call ***");
                    Console.WriteLine(" Member name: {0}", e.TargetSite);
                    Console.WriteLine(" Class defining member: {0}", e.TargetSite.DeclaringType);
                    Console.WriteLine(" Member type: {0}", e.TargetSite.MemberType);
                    Console.WriteLine(" Stack: {0}", e.StackTrace);
                    Console.WriteLine(" Message: {0}", e.Message);
                    Console.WriteLine(" Source: {0}", e.Source);
                    // By default, the data field is empty, so check for null. 
                    Console.WriteLine("\n-> Custom Data:");
                    if (e.Data != null && e.Data.Count > 0)
                    {
                        foreach (DictionaryEntry de in e.Data) Console.WriteLine("-> {0}: {1}", de.Key, de.Value);
                    }
                    else
                    {
                        Console.WriteLine("No Custom Data");
                    }
                }
            }

            if ( fSuccess )
            {
                Console.WriteLine("SheetMerge succeeded");

            }
            else
            {
                Console.WriteLine("SheetMerge FAILED");

            }
            Console.WriteLine("Press ENTER to continue...");
            Console.ReadLine();
        }
        // Launches web browser so user can authorize access to their Smartsheet account
        // Requires a redirect_uri - you must have a publicly accessible web page that will display the authorization code
        //      with instructions to copy and paste it into the application.  
        //      Optionally provide a countdown for how long the code is valid

        public static Token OAuthDesktopApp()
        {
            // fill in with the same data on your app registration page in Smartsheet
            string myclientid = "undhiisshsfdtjx6wn";
            string myclientsecret = "1oft5chvdgstgxi4oxd";
            string redirect_uri = "http://www.jeera.com/token/displaytoken.html";
            string myauthstate = "MY_STATE";

            // Setup the information that is necessary to request an authorization code
            OAuthFlow oauth = new OAuthFlowBuilder().SetClientId(myclientid)
                .SetClientSecret(myclientsecret).SetRedirectURL(redirect_uri).Build();

            // Create the URL that the user will go to grant authorization to the application
            // Specify what permission scope you need
            string url = oauth.NewAuthorizationURL(new Smartsheet.Api.OAuth.AccessScope[] {
        Smartsheet.Api.OAuth.AccessScope.READ_SHEETS,
        Smartsheet.Api.OAuth.AccessScope.CREATE_SHEETS,
        Smartsheet.Api.OAuth.AccessScope.WRITE_SHEETS }, myauthstate);

            // Launch the URL to authorize access
            Console.WriteLine("Launching web browser to grant permission.");
            System.Diagnostics.Process.Start(url);

            // After the user accepts or declines the authorization they are taken to the redirect URL
            //      where the user must copy the authorization code and paste into this app
            Console.WriteLine("Enter the authorization code copied from web page");
            string auth_code = Console.ReadLine();

            try
            {
                // recreate the ResponseURL so it can be parsed using the oauth object
                // TODO: create a constructor for the AuthorizationResult that will take the auth_code directly without have to pass in a URL
                string authorizationResponseURL = redirect_uri + "?code=" + auth_code;// + "&expires_in=239824&state=key%3DYOUR_VALUE";
                AuthorizationResult authResult = oauth.ExtractAuthorizationResult(authorizationResponseURL);

                // Get the token from the authorization result
                Token token = oauth.ObtainNewToken(authResult);
                Console.WriteLine("Successfully obtained a token to access the api");
                return token;
            }
            catch (Smartsheet.Api.SmartsheetException e)
            {
                Console.WriteLine("Authorization Error: {0}", e.Message);
                return null;
            }
        }

        private static bool MergeSheets(Token token)
        {
            bool fSuccess = true;

            // Use the Smartsheet Builder to create a SmartsheetClient object
            SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(token.AccessToken).Build();

            // Get the source sheet
            Sheet source_sheet = null;
            Sheet dest_sheet = null;
            try
            {
                source_sheet = smartsheet.SheetResources.GetSheet(source_id, null, null, null, null, null, null, null);
                dest_sheet = smartsheet.SheetResources.GetSheet(dest_id, null, null, null, null, null, null, null);
            }
            catch (SmartsheetRestException e)
            {
                long sheet_id = source_sheet == null ? source_id : dest_id;

                Console.WriteLine("\n*** Error Getting Sheet ID: {0} : {1} ***\n\n", sheet_id, e.Message);
                fSuccess = false;
                throw;
            }

            // check that the key_column exists on both sheets
            IList<Column> source_cols = source_sheet.Columns;
            IList<Column> dest_cols = dest_sheet.Columns;
            long source_key_col_ID = 0, dest_key_col_ID = 0;

            if ( (fSuccess = FindColumnTitle(source_cols, key_column, out source_key_col_ID)) == true )
            {
                fSuccess = FindColumnTitle(dest_cols, key_column, out dest_key_col_ID);
            }

            if ( !fSuccess )
            {
                Console.WriteLine("key column doesn't exist on both sheets");
                return fSuccess;
            }

            // create a map of columnIDs from source sheet to destination sheet
            Dictionary<long, long> map_source_to_dest_columnID = new Dictionary<long, long>();
            foreach (var s_col in source_cols)
            {
                //find the column title in the dest sheet
                bool fFound = false;
                for (int i = 0; i < dest_cols.Count && !fFound; i++)
                {
                    if ( dest_cols[i].Title == s_col.Title )
                    {
                        fFound = true;
                        map_source_to_dest_columnID.Add((long)s_col.Id, (long)dest_cols[i].Id);
                    }
                }
            }

            // Create a map of "key values" on the dest sheet to a row ID in the dest sheet
            Dictionary<string, long> map_key_value_to_row_ID_in_dest = new Dictionary<string, long>();
            foreach (var row in dest_sheet.Rows)
            {
                string key_value = GetCellStringValueFromColumnID(row.Cells, dest_key_col_ID);
                map_key_value_to_row_ID_in_dest.Add(key_value, (long)row.Id);
            }

 /*
            //debug - print the column ID map
            foreach (var item in map_source_to_dest_columnID)
            {
                Console.WriteLine("Map {0} to {1}", item.Key, item.Value);
            }
*/

            // iterate throught the rows in the source sheet and decide if it is
            // an update the destination sheet or a new row
            List<Row> update_rows = new List<Row>();
            List<Row> add_rows = new List<Row>();

            IList<Row> source_rows = source_sheet.Rows;
            IList<Row> dest_rows = dest_sheet.Rows;
            foreach (var row in source_rows)   // for each source row
            {
                // get the list of cells in the source row
                IList<Cell> source_row_cells = row.Cells;

                // get the key value from source
                string source_key = GetCellStringValueFromColumnID(source_row_cells, source_key_col_ID);
                Console.WriteLine("Looking up key value {0} ", source_key);

                // look for a matching key value in the destination sheet
                long dest_row_id;
                if ( map_key_value_to_row_ID_in_dest.TryGetValue(source_key, out dest_row_id) )
                {
                    // matching found on dest sheet, add row to the update list
                    Console.WriteLine("Match Found, need to update row");
                    Row update_row = CreateRowWithMappedCells(source_row_cells, map_source_to_dest_columnID);
                    update_row.Id = dest_row_id;
                    update_rows.Add(update_row);
                }
                else  // not found on dest sheet, add row to the new row list
                {
                    Console.WriteLine("Not Found, need to add row");
                    Row new_row = CreateRowWithMappedCells(source_row_cells, map_source_to_dest_columnID);
                    new_row.ToBottom = true;
                    add_rows.Add(new_row);
                }

                /*
                                Console.Write("source row {0}", item.RowNumber);
                                foreach (var cell in source_row_cells)
                                {
                                    Console.Write("  cell {0} : {1} ", cell.ColumnId, cell.DisplayValue);
                                }
                */
            }

            Console.WriteLine("\n\nAdding {0} rows, Updating {1} rows.", add_rows.Count, update_rows.Count);

            // update Rows
            SheetRowResources smartsheet_rowresources = smartsheet.SheetResources.RowResources;
            if ( update_rows.Count > 0 )
            {
                IList<Row> ret_update_rows = smartsheet_rowresources.UpdateRows(dest_id, update_rows);
            }

            // add Rows
            if ( add_rows.Count > 0 )
            {
                IList<Row> ret_add_rows = smartsheet_rowresources.AddRows(dest_id, add_rows);
            }

            // return
            return fSuccess;
        }

        // Given a list of cells and a mapping dictionary, create a new list of cells where the column ids are
        // mapped from a source column ID to a destination column ID
        private static Row CreateRowWithMappedCells(IList<Cell> source_row_cells, Dictionary<long, long> map_source_to_dest_columnID)
        {
            Row new_row = null;

            List<Cell> cells_to_add = new List<Cell>();
            foreach (var cell in source_row_cells)
            {
                if (cell.Value != null)
                {
                    long new_col_id = 0;
                    if (map_source_to_dest_columnID.TryGetValue((long)cell.ColumnId, out new_col_id))
                    {
                        Cell new_cell = new Cell();
                        new_cell.Value = cell.Value;
                        new_cell.ColumnId = new_col_id;
                        cells_to_add.Add(new_cell);
                    }
                }
            }
            if (cells_to_add.Count > 0)
            {
                new_row = new Row();
                new_row.Cells = cells_to_add;
            }

            return new_row;

        }

        private static bool FindColumnTitle(IList<Column> col_list, string title, out long col_ID)
        {
            bool fFound = false;
            col_ID = 0;
            for (int i = 0; i < col_list.Count && !fFound; i++)
            {
                if (col_list[i].Title == title)
                {
                    fFound = true;
                    col_ID = (long)col_list[i].Id;
                    //Console.WriteLine("key column {0} found in source sheet", find_title);
                }
            }

            return fFound;
        }

        private static string GetCellStringValueFromColumnID(IList<Cell> row_cells, long col_ID)
        {
            string key = "";
            bool fFound = false;
            for (int i = 0; i < row_cells.Count && !fFound; i++)
            {
                if ( row_cells[i].ColumnId == col_ID)
                {
                    key = row_cells[i].Value.ToString();
                    fFound = true;
                }
            }

            return key;
        }
    }
}
