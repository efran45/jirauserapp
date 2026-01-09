# Jira User & Group Management Tool

A GUI application for managing Jira users, groups, and product access with advanced filtering and export capabilities.

## Prerequisites

- Python 3.x installed on your system
- Jira instance with API access
- Jira API token (see Setup section)

## Installation

1. **Install Required Python Packages**

   Open your command prompt/terminal and run:
   ```bash
   pip install tkcalendar python-dateutil requests keyring
   ```

2. **Download the Script**

   Save `jira_user_app.py` to your preferred location.

## Setup

### Getting Your Jira API Token

1. Log in to your Atlassian account at https://id.atlassian.com
2. Go to **Security** → **API tokens**
3. Click **Create API token**
4. Give it a name (e.g., "Jira User Management")
5. Copy the token (you won't be able to see it again!)

### Optional: Organization API Setup

For last login data, you'll need an Organization API key:

1. Go to https://admin.atlassian.com
2. Navigate to **Settings** → **API keys**
3. Create a new API key with appropriate permissions
4. Copy the key for use in the application

## Running the Application

Open your command prompt/terminal, navigate to the folder containing the script, and run:

```bash
python jira_user_app.py
```

**Note:** Use `python` (not `python3` or `py`) to ensure it uses the correct Python installation with all packages.

## First-Time Configuration

1. **Launch the application** - The Configuration tab will open automatically

2. **Enter your Jira connection details:**
   - **Jira URL**: Your full Jira URL (e.g., `https://your-company.atlassian.net`)
   - **Email**: Your Atlassian account email
   - **Jira API Token**: The token you created in Setup

3. **Optional - Enable Organization API:**
   - Check "Enable Organization API" if you want last login data
   - Enter your Organization API Key
   - Click "Get Org ID" to automatically fetch your Organization ID

4. **Validate your token:**
   - Click the **✓ Validate Token** button to test your connection
   - Wait for confirmation that credentials are valid

5. **Save settings:**
   - Check "Remember credentials" to save your Jira URL and email (tokens are stored securely)

## Using the Application

### Users & Groups Tab

**Fetching Data:**
- Click **📥 Fetch Users** to retrieve all users from your Jira instance
- Click **👥 Fetch Groups** to retrieve all groups
- Groups are expandable - click to view members

**Searching & Filtering:**
- Use the search box to filter by name, email, or account ID
- Filter users by Status (Active/Inactive)
- Filter users by Type (atlassian/app/customer)
- Filter by last login date range (when Org API is enabled)

**Right-Click Menu (Users only):**
- **Open User Profile** - Opens the user's profile in Atlassian Admin
- **Copy Account ID** - Copies the account ID to clipboard
- **Copy Email** - Copies the email to clipboard

**Exporting:**
- Click **💾 Export CSV** to export current view to CSV file
- CSV includes all visible data and respects current filters
- Files are saved with timestamp: `jira_users_YYYYMMDD_HHMMSS.csv`

### Products Tab

**Viewing Product Access:**
- Click **📦 Fetch Products** to retrieve all products and their users
- View which users have access to each product
- See product names and user counts

**Exporting:**
- Export product access data to CSV format

## Features

✓ View all Jira users with detailed information  
✓ Browse groups and their members (expandable tree view)  
✓ Advanced search and filtering capabilities  
✓ View last login dates (with Organization API)  
✓ Export data to CSV format  
✓ Secure credential storage using system keyring  
✓ Right-click context menu for quick actions  
✓ Sort by any column (click column headers)  
✓ Product access management view  

## Troubleshooting

### "ModuleNotFoundError: No module named 'tkcalendar'"

Make sure you're using `python` command (not `python3` or `py`). Test with:
```bash
python -c "import tkcalendar; print('OK')"
```

If this works but the script doesn't, you have multiple Python installations. Always use `python`.

### "Authentication failed" or "401 Unauthorized"

- Verify your Jira URL is correct (include https://)
- Check that your email matches your Atlassian account
- Regenerate your API token if needed
- Make sure your account has proper permissions

### Application won't start or shows errors

- Ensure all required packages are installed
- Check that you have Python 3.x (not Python 2.x)
- Try running from command line to see full error messages

### Organization API not working

- Verify you have the correct Organization API key
- Make sure your account has admin access
- Check that the Organization ID was fetched correctly
- Some features require specific Atlassian plan levels

## Data Security

- API tokens are stored securely in your system's keyring
- Jira URL and email are saved in plain text (if "Remember credentials" is checked)
- Organization API keys are also stored in system keyring
- No passwords are ever saved

## Support

For issues related to:
- **Jira API**: Check Atlassian's API documentation at https://developer.atlassian.com
- **API Tokens**: Visit https://support.atlassian.com/atlassian-account/docs/manage-api-tokens-for-your-atlassian-account/
- **Script errors**: Check that all prerequisites are installed and you're using the correct Python version

## Tips

- Use Ctrl+F to quickly find users in large lists
- Column sorting persists between fetches
- Right-click on users for quick access to their profile
- Export before making changes to have a backup
- The tree view for groups can be expanded/collapsed by clicking the arrows
- Use date filters to find inactive users for license management
