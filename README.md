# HubSpot Pipeline Reporter

Automated weekly HubSpot pipeline reporting tool that exports deal data to Excel spreadsheets.

## Overview

This tool connects to your HubSpot account via API, fetches pipeline deal data, and automatically generates formatted Excel reports. It can be scheduled to run weekly or on-demand, making it perfect for regular sales reporting and pipeline analytics.

## Features

- 🔌 **HubSpot API Integration** - Securely connects to HubSpot using Private App tokens
- 📊 **Excel Export** - Generates formatted Excel reports with styled headers
- ⏰ **Automated Scheduling** - Weekly automated reporting (runs every Monday at 9 AM)
- 📈 **Pipeline Analytics** - Fetches deal names, amounts, stages, close dates, and creation dates
- 🔄 **Pagination Support** - Handles large datasets automatically
- 🎨 **Professional Formatting** - Auto-adjusts column widths and applies styling

## Tech Stack

- **Python 3** - Core programming language
- **requests** - HTTP library for API calls
- **openpyxl** - Excel file generation
- **python-dotenv** - Environment variable management
- **schedule** - Task scheduling library

## Installation

### Prerequisites

- Python 3.7 or higher
- HubSpot account with Private App access
- pip (Python package manager)

### Setup Steps

1. **Clone the repository**
   ```bash
   git clone https://github.com/owen99-tech/hubspot-pipeline-reporter.git
   cd hubspot-pipeline-reporter
   ```

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Configure environment variables**
   ```bash
   cp .env.example .env
   ```
   
   Edit `.env` and add your HubSpot API key:
   ```
   HUBSPOT_API_KEY=your_private_app_token_here
   ```

### Getting Your HubSpot API Key

1. Log in to your HubSpot account
2. Navigate to Settings > Integrations > Private Apps
3. Create a new Private App or use an existing one
4. Ensure the app has the following scopes:
   - `crm.objects.deals.read`
   - `crm.schemas.deals.read`
5. Copy the access token and paste it into your `.env` file

## Usage

### Run Once (On-Demand)

To generate a report immediately:

```bash
python hubspot_reporter.py
```

This will:
- Connect to HubSpot
- Fetch all pipelines
- Export the first pipeline to an Excel file in the `reports/` directory

### Run with Scheduler (Weekly Automation)

To set up automatic weekly reports:

```bash
python scheduler.py
```

The scheduler will:
- Run every Monday at 9:00 AM
- Generate fresh reports automatically
- Keep running in the background
- Press `Ctrl+C` to stop

### Customization

#### Change the Pipeline

By default, the tool exports the first pipeline. To select a specific pipeline, edit `hubspot_reporter.py`:

```python
# Instead of:
pipeline = pipelines[0]

# Use:
pipeline = next(p for p in pipelines if p['label'] == 'Sales Pipeline')
```

#### Change the Schedule

To modify the reporting schedule, edit `scheduler.py`:

```python
# Weekly on Monday at 9 AM (current)
schedule.every().monday.at("09:00").do(scheduled_job)

# Daily at 8 AM
schedule.every().day.at("08:00").do(scheduled_job)

# Every Friday at 5 PM
schedule.every().friday.at("17:00").do(scheduled_job)
```

## Output

Reports are saved in the `reports/` directory with timestamps:

```
reports/
├── Sales_Pipeline_20251230_142530.xlsx
└── Sales_Pipeline_20251223_090015.xlsx
```

Each Excel file includes:
- Deal Name
- Amount
- Deal Stage
- Close Date
- Created Date

## Project Structure

```
hubspot-pipeline-reporter/
├── hubspot_reporter.py   # Main reporter script
├── scheduler.py          # Weekly automation scheduler
├── requirements.txt      # Python dependencies
├── .env.example          # Environment variable template
├── .gitignore           # Git ignore rules
├── README.md            # This file
└── reports/             # Generated Excel reports (auto-created)
```

## Development

### Running in Development

```bash
# Set up virtual environment (recommended)
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Run the reporter
python hubspot_reporter.py
```

## Troubleshooting

### Authentication Errors

**Error**: `HubSpot API key not found`
- Ensure `.env` file exists and contains `HUBSPOT_API_KEY`
- Check that the API key is valid and not expired

**Error**: `401 Unauthorized`
- Verify your API token has the required scopes
- Check that the token hasn't been revoked

### No Deals Found

- Verify that deals exist in your selected pipeline
- Check that your API token has permission to read deals
- Ensure the pipeline ID matches an active pipeline

## Security Notes

- **Never commit `.env` file** - It contains sensitive API keys
- **Use Private Apps** - More secure than API keys
- **Rotate tokens regularly** - Best practice for API security
- **.env is gitignored** - Automatically excluded from version control

## License

MIT License - Feel free to use and modify for your needs.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Support

For issues or questions:
- Open an issue on GitHub
- Check HubSpot API documentation: https://developers.hubspot.com/docs/api/overview
