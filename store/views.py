import os
import re
import json
import logging
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
from django.shortcuts import render
from django.http import JsonResponse
from django.conf import settings

# Set up logging
logger = logging.getLogger(__name__)

def get_google_sheet():
    """Initialize Google Sheets API and return the worksheet"""
    try:
        # Path to your service account credentials JSON file
        creds_path = os.path.join(settings.BASE_DIR, 'credentials.json')
        
        # If using environment variable instead of file
        creds_json = os.getenv('GOOGLE_CREDS_JSON')
        if creds_json:
            creds_info = json.loads(creds_json)
        else:
            with open(creds_path) as f:
                creds_info = json.load(f)
        
        # Authorize the API
        scope = ['https://spreadsheets.google.com/feeds',
                 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_info, scope)
        client = gspread.authorize(creds)
        
        # Open the Google Sheet (replace with your sheet ID or name)
        sheet = client.open_by_key('1K2WzzOqJ8hX_7a6vZMXL2F2sTThQqUV3gFcrHaWs9vY').sheet1
        return sheet
    except Exception as e:
        logger.error(f"Error initializing Google Sheet: {str(e)}")
        raise

def is_valid_email(email):
    """Validate email format"""
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return re.match(pattern, email) is not None

# Set up logging
logger = logging.getLogger(__name__)

def is_valid_email(email):
    """Validate email format"""
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return re.match(pattern, email) is not None

def get_excel_path():
    """Get the path to the Excel file, creating directory if needed"""
    try:
        # Create a 'subscriptions' directory in the project root if it doesn't exist
        excel_dir = os.path.join(settings.BASE_DIR, 'subscriptions')
        if not os.path.exists(excel_dir):
            os.makedirs(excel_dir, exist_ok=True)
            logger.info(f"Created directory: {excel_dir}")
            
        # Ensure directory is writable
        if not os.access(excel_dir, os.W_OK):
            raise PermissionError(f"Cannot write to directory: {excel_dir}")
            
        excel_path = os.path.join(excel_dir, 'subscribers.xlsx')
        logger.info(f"Using Excel file: {excel_path}")
        return excel_path
        
    except Exception as e:
        logger.error(f"Error getting Excel path: {str(e)}")
        raise

def save_to_google_sheet(email):
    """Save email to Google Sheet with enhanced duplicate checking"""
    if not is_valid_email(email):
        return False, "Invalid email format"
    
    # Normalize email (lowercase and remove any whitespace)
    email = email.lower().strip()
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
    try:
        sheet = get_google_sheet()
        
        # Ensure headers exist
        if sheet.row_count == 0:
            sheet.append_row(['Email', 'Subscription Date'])
        
        # Get all existing emails in one batch (more efficient)
        try:
            # Try to get all values at once
            all_values = sheet.get_all_values()
            # Skip header row if it exists
            existing_emails = [row[0].lower().strip() for row in all_values[1:] if row and row[0]]
        except Exception as e:
            logger.error(f"Error reading existing emails: {str(e)}")
            return False, "Error checking existing subscriptions. Please try again."
        
        # Check for existing email (case-insensitive)
        if email in existing_emails:
            return False, "This email is already subscribed!"
        
        # Add new subscription
        try:
            sheet.append_row([email, timestamp])
            logger.info(f"Successfully saved subscription for {email}")
            return True, "Thank you for subscribing!"
        except Exception as e:
            logger.error(f"Error appending row: {str(e)}")
            # Double check if the row was actually added despite the error
            try:
                all_values = sheet.get_all_values()
                last_row = all_values[-1] if all_values else None
                if last_row and last_row[0].lower().strip() == email:
                    return True, "Thank you for subscribing!"
            except:
                pass
            return False, "Error saving your subscription. Please try again."
        
    except Exception as e:
        error_msg = f"Error in save_to_google_sheet: {str(e)}"
        logger.error(error_msg, exc_info=True)
        return False, "An unexpected error occurred. Please try again later."

def index(request):
    if request.method == 'POST' and request.headers.get('X-Requested-With') == 'XMLHttpRequest':
        email = (request.POST.get('email') or '').strip()
        
        if not email:
            return JsonResponse(
                {'status': 'error', 'message': 'Email is required'}, 
                status=400
            )
        
        try:
            success, message = save_to_google_sheet(email)
            if success:
                return JsonResponse(
                    {'status': 'success', 'message': message}
                )
            else:
                return JsonResponse(
                    {'status': 'error', 'message': message},
                    status=400
                )
                
        except Exception as e:
            logger.error(f"Error in index view: {str(e)}")
            return JsonResponse(
                {'status': 'error', 'message': 'An error occurred. Please try again later.'}, 
                status=500
            )
    
    # For GET requests, just render the page
    return render(request, 'index.html')
