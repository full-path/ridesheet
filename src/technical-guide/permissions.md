# RideSheet and Google Workspace Permissions

## Why Google Asks for Permission

When you use RideSheet, Google asks you to approve what RideSheet can access in your account. This protects your data - apps can't access anything unless you explicitly allow it. 

---

## Summary of RideSheet Permissions

RideSheet uses the most restrictive permissions possible to accomplish its mission. Your data always stays in your Google account. RideSheet cannot access your Gmail, Calendar, Contacts, or other Google services. Here is a list of what permissions RideSheet requests and why:

| Permission | Used For |
|-----------------------|----------|
| Access your RideSheet spreadsheet only | Managing trips, customers, and runs |
| Show menus and dialogs | User interface |
| Connect to Google services | Integration with Google Maps |
| Access files RideSheet creates | Saving manifests to your Google Drive |
| Access Google Docs | Reading your manifest template, creating formatted manifests |

---

## Customizing Your Permission Selection

You can grant permissions individually to match how you'll use RideSheet.

### Required Permissions

If you deny these permissions, you'll lose basic RideSheet functionality.

- **Spreadsheet access** - RideSheet cannot perform its most basic functions without access to the spreadsheet it is installed on.
- **Menus and dialogs** - RideSheet needs this to provide messages about the results of the tasks it is performing.
- **External connections** - Needed for integration with Google Maps for address validation, distance estimation, and deadhead calculation.

### Optional Permissions

These permissions are needed only for generating driver manifests. If you don't need driver manifests, you can decline to authorize access to them.

- **Files RideSheet creates**
- **Google Docs access**

---

## Revoking Permissions

**You can revoke RideSheet's permissions at any time:**

1. Visit [https://myaccount.google.com/connections](https://myaccount.google.com/connections)
2. Find RideSheet in the list
3. Click to view details
4. Click "Remove access" or "Disconnect"

**What happens when you revoke:**

- RideSheet immediately loses all access to your account
- RideSheet cannot read or modify any data
- Any features you were using will stop working
- All your data remains safe in your Google account

The next time you attempt to use RideSheet features, you will be prompted to authorize access, at which time you can select only the permissions you want to grant.

!!! warning
    Revoking access only affects **your** individual user account. It will not affect the permissions that other users have granted when they use RideSheet. To reset access for the entire organization, each individual user will need to go through the steps above.

    To remove access for your organization as a whole, remove shared access to RideSheet or delete it from Google Drive entirely. 

    You can retain all of RideSheet's data by exporting it to Excel. Exporting to Excel keeps all your data and removes all of RideSheet's programming scripts.

---

## Reviewing Permissions

**When you first use RideSheet:**

Google shows you exactly what permissions RideSheet requests. Review this screen before clicking "Allow."

**After you have granted access:**

1. Visit [https://myaccount.google.com/connections](https://myaccount.google.com/connections)
2. Find "RideSheet" in the list
3. Click to see what RideSheet can access from your user account

---

## Your Data Stays Private

**Where your data is stored:**

- In your RideSheet spreadsheet (in your Google account)
- In the manifests RideSheet creates (in your Google Drive)
- Nowhere else

**What RideSheet doesn't do:**

- Upload your data to external servers
- Share your data with third parties
- Access your data when you're not using it
- Store copies outside your Google account

**The only external connection** is to Google Maps when looking up addresses or calculating trip distances. This uses the same Google Maps service you use when you open Google Maps directly.
