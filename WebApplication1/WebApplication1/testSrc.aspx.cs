using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace WebApplication1
{
    public partial class testSrc : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            string userName = "AMAT\\jrybicki118368";
            string groupName = "IMDb Owners";
            string hostWebURL = "https://appliedapps.amat.com/sites/IMDb";
            EnsureCurrentUserInSpAndADGroup(hostWebURL, groupName, userName);
        }

        /// <summary>
        /// CHECK: USER IN SP GROUP AND AD GROUP(ADDED IN SP GROUP)
        /// </summary>
        public bool EnsureCurrentUserInSpAndADGroup(string hostWebURL, string groupName, string userName)
        {
            try
            {
                bool retValue = false;
                using (ClientContext context = new ClientContext(hostWebURL))
                {
                    GroupCollection groups = context.Web.SiteGroups;
                    context.Load(groups, group => group.Include(grp => grp.Title));
                    context.ExecuteQuery();
                    foreach (Group g in groups)
                    {
                        if (g.Title == groupName)
                        {
                            try
                            {
                                UserCollection oUsers = g.Users;
                                context.Load(oUsers);
                                context.ExecuteQuery();
                                foreach (User usr in oUsers)
                                {
                                    if (usr.PrincipalType == PrincipalType.SecurityGroup)
                                    {
                                        System.DirectoryServices.AccountManagement.PrincipalContext ctx = new System.DirectoryServices.AccountManagement.PrincipalContext(
                                            System.DirectoryServices.AccountManagement.ContextType.Domain, "AMAT");

                                        // find a user
                                        System.DirectoryServices.AccountManagement.UserPrincipal uPrincipal = System.DirectoryServices.AccountManagement.UserPrincipal.FindByIdentity(ctx, userName);


                                        using (System.DirectoryServices.AccountManagement.PrincipalSearchResult<System.DirectoryServices.AccountManagement.Principal> psGroups = uPrincipal.GetAuthorizationGroups())
                                        {
                                            retValue = psGroups.OfType<System.DirectoryServices.AccountManagement.GroupPrincipal>().Any(gp => gp.Name.Equals(usr.Title, StringComparison.OrdinalIgnoreCase));
                                        }
                                    }
                                    else
                                    {
                                        if (usr.LoginName.ToLower().Contains(userName.ToLower()))
                                            retValue = true;
                                    }
                                    if (retValue)
                                        return retValue;
                                }
                            }
                            catch (Exception ex)
                            {
                                if (ex.Message.Contains("User cannot be found")) retValue = false;
                                else throw ex;
                            }
                        }
                        if (retValue)
                            return retValue;
                    }
                }
                return retValue;
            }
            catch (Exception Ex)
            { throw Ex; }
        }
    }
} 