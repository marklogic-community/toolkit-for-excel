// -----------------------------------------------------------------------
// 
//  Copyright (C) Microsoft Corporation.  All rights reserved.
// 
// THIS CODE AND INFORMATION ARE PROVIDED AS IS WITHOUT WARRANTY OF ANY
// KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
// IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
// PARTICULAR PURPOSE.
// -----------------------------------------------------------------------

using System;
using System.Collections;
using System.ComponentModel;
using System.Configuration.Install;
using System.Security;
using System.Security.Permissions;
using Microsoft.VisualStudio.Tools.Office.Runtime.Security;

namespace InclusionListCustomActions
{
    [RunInstaller(true)]
    public class TrustInstaller
        : Installer
    {
        const string RSA_PublicKey = "<RSAKeyValue><Modulus>rOs9i5k01Mk8L/+ANs09CsokLxErn/3J3DFfLixylmvKUXwQiSCcotAyqEPCEpoq6zUrVaNt9RwxYLaiOjC2w+v2gNCsD29Sh+pQr/gCAsIJyfrlgpsCp+X+dhuC3B+7RsXJIua0MIIvJkwbPNqOeuF87Pqe3fRtcJ2aEW2Kzlk=</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>";

        public override void Install(IDictionary stateSaver)
        {
            try
            {
                SecurityPermission permission =
                    new SecurityPermission(PermissionState.Unrestricted);
                permission.Demand();
            }
            catch (SecurityException)
            {
                throw new InstallException(
                    "You have insufficient privileges to " +
                    "register a trust relationship. Start Word " +
                    "and confirm the trust dialog to run the addin.");
            }
            Uri deploymentManifestLocation = null;
            if (Uri.TryCreate(Context.Parameters["deploymentManifestLocation"],
                UriKind.RelativeOrAbsolute, out deploymentManifestLocation) == false)
            {
                throw new InstallException(
                    "The location of the deployment manifest is missing or invalid.");
            }
            AddInSecurityEntry entry = new AddInSecurityEntry(
                            deploymentManifestLocation, RSA_PublicKey);
            UserInclusionList.Add(entry);
            stateSaver.Add("entryKey", deploymentManifestLocation);
            base.Install(stateSaver);
        }

        public override void Uninstall(IDictionary savedState)
        {
            Uri deploymentManifestLocation = (Uri)savedState["entryKey"];
            if (deploymentManifestLocation != null)
            {
                UserInclusionList.Remove(deploymentManifestLocation);
            }
            base.Uninstall(savedState);
        }

        public override void Commit(IDictionary savedState)
        {
            base.Commit(savedState);
        }

        public override void Rollback(IDictionary savedState)
        {
            base.Rollback(savedState);
        }
    }
}
