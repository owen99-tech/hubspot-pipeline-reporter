#!/usr/bin/env python3
"""
Scheduler for HubSpot Pipeline Reporter
Runs the reporter weekly using schedule library
"""

import schedule
import time
from datetime import datetime
from hubspot_reporter import main as run_reporter

def scheduled_job():
    """
    Job to run on schedule
    """
    print(f"\n{'='*50}")
    print(f"Starting scheduled report generation at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"{'='*50}\n")
    
    try:
        run_reporter()
        print(f"\n{'='*50}")
        print(f"Report generation completed successfully")
        print(f"{'='*50}\n")
    except Exception as e:
        print(f"\nError during scheduled report generation: {e}\n")

if __name__ == "__main__":
    # Schedule the job to run every Monday at 9:00 AM
    schedule.every().monday.at("09:00").do(scheduled_job)
    
    print("HubSpot Pipeline Reporter Scheduler Started")
    print("Report will be generated every Monday at 9:00 AM")
    print("Press Ctrl+C to stop\n")
    
    # Run the job immediately on startup (optional)
    # scheduled_job()
    
    # Keep the script running
    while True:
        schedule.run_pending()
        time.sleep(60)  # Check every minute
