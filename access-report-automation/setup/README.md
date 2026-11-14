# Setting up the app registration

The report signs in as an application, not as a person. This guide creates that application in Entra ID. You need a role that can grant admin consent (Global Administrator or Privileged Role Administrator).

At the end you have:

- the **Directory (tenant) id**
- the **Application (client) id**
- either a **`.secret.pem`** file (certificate sign-in) or a federated credential (nothing to store)

## 1. Create the app registration

1. [Entra admin center](https://entra.microsoft.com) > **Identity > Applications > App registrations > New registration**.
2. Name: `SharePoint access report`. Supported account types: **Accounts in this organizational directory only**. No redirect URI. **Register**.
3. On the Overview page, copy **Application (client) ID** and **Directory (tenant) ID**.

## 2. Add the permissions

**API permissions > Add a permission**. All four are **Application permissions** (not Delegated).

| API | Permission |
| --- | --- |
| **SharePoint** | `Sites.FullControl.All` |
| **Microsoft Graph** | `Sites.Read.All` |
| **Microsoft Graph** | `User.Read.All` |
| **Microsoft Graph** | `GroupMember.Read.All` |

Then **Grant admin consent for <your organisation>**. Every row must show a green tick under Status.

You can remove the delegated `User.Read` permission that new registrations get by default; it is not used.

Why these, and how to narrow them, is in the main [README](../README.md#permissions-and-why).

## 3. Give the app a way to sign in

SharePoint does not accept app-only sign-in with a client secret. Choose one of:

### Option A: certificate (recommended to start with)

On any machine with `openssl` (Linux, macOS, WSL, Git Bash):

```bash
./new-certificate.sh access-report 730
```

This writes two files:

- `access-report.cer`: the public certificate. In the app registration: **Certificates & secrets > Certificates > Upload certificate**. The thumbprint shown after upload should match the one the script printed.
- `access-report.secret.pem`: the certificate with its private key. Store the **whole content** as the GitHub secret `ACCESS_REPORT_CERTIFICATE` (or wherever your scheduler keeps secrets), then delete the file.

The certificate is valid for the number of days you gave (730 here). From 45 days before it expires, the report's Summary sheet and the run log carry a warning. To renew: run the script again with a new name, upload the new `.cer`, replace the secret, then delete the old certificate from the app registration.

Already have a certificate as a `.pfx`? Convert it:

```bash
openssl pkcs12 -in mycert.pfx -nodes -out access-report.secret.pem
```

### Option B: federated credential (GitHub Actions, nothing stored)

GitHub proves the identity of the workflow run to Entra ID; there is no certificate to keep or renew.

1. App registration > **Certificates & secrets > Federated credentials > Add credential**.
2. Scenario: **GitHub Actions deploying Azure resources**.
3. Organization and repository: where the workflow lives. Entity type: **Branch**, with the default branch name (scheduled runs come from the default branch). Audience: leave `api://AzureADTokenExchange`.
4. In GitHub, set the repository variable `ACCESS_REPORT_AUTH` to `federated`. No `ACCESS_REPORT_CERTIFICATE` secret is needed.

Certificate sign-in to SharePoint is the long-established route; federated sign-in is newer and was only tested against the simulated tenant. If `check` reports that SharePoint refuses the token with federated sign-in, use Option A.

## 4. Check

Set the variables and secrets as described in the main [README](../README.md#3-tell-github), then run the workflow with mode **check** (or `node dist/cli.js check` locally). It tests each of the steps above separately and says what to fix.

## Removing it again

Delete the app registration. That revokes all access at once. The reports already in the folder stay.
