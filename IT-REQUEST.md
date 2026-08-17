# Request: Entra app registration for Outlook draft creation

*Forward this to whoever administers the Gateway Global Microsoft 365 tenant. It's a
two-minute job in the Entra admin centre.*

---

Hi,

I'm running a small internal tool that creates **draft** emails in my own Outlook mailbox so
I can review and send them manually. To let it talk to Microsoft Graph, I need an app
registration in our tenant. Could you set one up?

## What I need

**Entra admin centre → App registrations → New registration**

| Setting | Value |
|---|---|
| Name | `Prospect Draft Tool` |
| Supported account types | Accounts in this organizational directory only (single tenant) |
| Redirect URI | *(leave blank)* |

Then, on the new registration:

**Authentication →** set **"Allow public client flows"** to **Yes**.
*(This enables device-code sign-in. It means no client secret is needed, so there's no
credential for me to store or leak.)*

**API permissions →** add these **delegated** Microsoft Graph permissions:

| Permission | Why |
|---|---|
| `Mail.ReadWrite` | Create draft messages in my own mailbox |
| `User.Read` | Confirm which account signed in |
| `offline_access` | Keep the session alive without re-authenticating daily |

Grant admin consent if our tenant requires it for delegated permissions.

## What to send me back

Just the **Application (client) ID** and the **Directory (tenant) ID** from the registration's
Overview page. Neither is a secret.

## Why this is low risk

- **Delegated, not application permissions.** The tool acts only as me, and only on my own
  mailbox. It cannot touch anyone else's mail.
- **No send permission.** I've deliberately not requested `Mail.Send`. The tool is
  technically incapable of sending an email — it can only create drafts that I open and
  send myself from Outlook.
- **No stored secret.** Public client with device-code flow means no client secret exists.
  Sign-in happens in a browser against our normal login, MFA and Conditional Access included.
- **Revocable at any time.** Delete the registration and the tool stops working immediately.

Happy to walk through what it does if useful.

Thanks,
Mo
