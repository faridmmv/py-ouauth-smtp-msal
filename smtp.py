#!/usr/bin/env python3
"""
OAuth SMTP Authentication with Azure Entra and MSAL for Python

This module demonstrates how to authenticate SMTP connections using OAuth 2.0
Client Credentials flow with Microsoft Azure Entra ID and the Microsoft 
Authentication Library (MSAL) for Python.
"""

import os
import sys
import base64
import smtplib
import logging
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from typing import Optional

import msal


# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


class Config:
    """Configuration class for OAuth and SMTP settings."""
    
    def __init__(self):
        self.tenant_id = os.getenv('AZURE_TENANT_ID')
        self.client_id = os.getenv('AZURE_CLIENT_ID')
        self.client_secret = os.getenv('AZURE_CLIENT_SECRET')
        self.user_email = os.getenv('USER_EMAIL')
        self.smtp_server = 'smtp.office365.com'
        self.smtp_port = 587
        
    def validate(self) -> bool:
        """Validate that all required configuration is present."""
        required_fields = [
            ('AZURE_TENANT_ID', self.tenant_id),
            ('AZURE_CLIENT_ID', self.client_id),
            ('AZURE_CLIENT_SECRET', self.client_secret),
            ('USER_EMAIL', self.user_email)
        ]
        
        missing_fields = [name for name, value in required_fields if not value]
        
        if missing_fields:
            logger.error(f"Missing required environment variables: {', '.join(missing_fields)}")
            return False
        
        return True


class SMTPOAuthClient:
    """SMTP client with OAuth 2.0 Client Credentials authentication support."""
    
    def __init__(self, config: Config):
        self.config = config
        self.access_token: Optional[str] = None
        self.smtp_client: Optional[smtplib.SMTP] = None
        
    def authenticate_with_client_credentials(self) -> bool:
        """
        Authenticate using Client Credentials flow (OAuth2 authorization code flow for applications).
        
        Returns:
            bool: True if authentication successful, False otherwise
        """
        try:
            # Create confidential client application
            authority = f"https://login.microsoftonline.com/{self.config.tenant_id}"
            app = msal.ConfidentialClientApplication(
                client_id=self.config.client_id,
                client_credential=self.config.client_secret,
                authority=authority
            )
            
            # Define scopes for application permissions
            scopes = ["https://outlook.office365.com/.default"]
            
            # Acquire token using client credentials
            result = app.acquire_token_for_client(scopes=scopes)
            
            if "access_token" in result:
                self.access_token = result["access_token"]
                logger.info("Client credentials authentication successful!")
                return True
            else:
                error_description = result.get("error_description", "Unknown error")
                logger.error(f"Authentication failed: {error_description}")
                return False
                
        except Exception as e:
            logger.error(f"Authentication error: {str(e)}")
            return False
    
    def connect_smtp(self) -> bool:
        """
        Establish SMTP connection with OAuth authentication.
        
        Returns:
            bool: True if connection successful, False otherwise
        """
        if not self.access_token:
            logger.error("No access token available, please authenticate first")
            return False
        
        try:
            # Create SMTP connection
            self.smtp_client = smtplib.SMTP(self.config.smtp_server, self.config.smtp_port)
            
            # Enable debug output (optional)
            # self.smtp_client.set_debuglevel(1)
            
            # Send EHLO command first
            self.smtp_client.ehlo()
            
            # Start TLS encryption
            self.smtp_client.starttls()
            
            # Send EHLO again after STARTTLS (required)
            self.smtp_client.ehlo()
            
            # Perform OAuth authentication
            auth_string = f"user={self.config.user_email}\x01auth=Bearer {self.access_token}\x01\x01"
            auth_b64 = base64.b64encode(auth_string.encode('ascii')).decode('ascii')
            
            # Send AUTH command
            code, response = self.smtp_client.docmd("AUTH", f"XOAUTH2 {auth_b64}")
            if code not in (235, 250):  # 235 = Authentication successful, 250 = OK
                raise smtplib.SMTPAuthenticationError(code, response)
            
            logger.info("SMTP connection established successfully!")
            return True
            
        except smtplib.SMTPAuthenticationError as e:
            logger.error(f"SMTP authentication failed: {str(e)}")
            return False
        except Exception as e:
            logger.error(f"SMTP connection failed: {str(e)}")
            return False
    
    def test_smtp_connection(self) -> bool:
        """
        Test the SMTP connection.
        
        Returns:
            bool: True if test successful, False otherwise
        """
        if not self.smtp_client:
            logger.error("SMTP client not connected")
            return False
        
        try:
            # Test NOOP command
            status, message = self.smtp_client.noop()
            if status == 250:
                logger.info("SMTP connection test successful!")
                return True
            else:
                logger.error(f"SMTP connection test failed: {status} {message}")
                return False
        except Exception as e:
            logger.error(f"SMTP connection test failed: {str(e)}")
            return False
    
    def send_test_email(self, to_email: str, subject: str, body: str) -> bool:
        """
        Send a test email.
        
        Args:
            to_email (str): Recipient email address
            subject (str): Email subject
            body (str): Email body
            
        Returns:
            bool: True if email sent successfully, False otherwise
        """
        if not self.smtp_client:
            logger.error("SMTP client not connected")
            return False
        
        try:
            # Create email message
            msg = MIMEMultipart()
            msg['From'] = self.config.user_email
            msg['To'] = to_email
            msg['Subject'] = subject
            
            # Add body to email
            msg.attach(MIMEText(body, 'plain'))
            
            # Send email
            text = msg.as_string()
            self.smtp_client.sendmail(self.config.user_email, to_email, text)
            
            logger.info("Test email sent successfully!")
            return True
            
        except Exception as e:
            logger.error(f"Failed to send test email: {str(e)}")
            return False
    
    def close(self):
        """Close the SMTP connection."""
        if self.smtp_client:
            try:
                self.smtp_client.quit()
                logger.info("SMTP connection closed")
            except Exception as e:
                logger.warning(f"Error closing SMTP connection: {str(e)}")
            finally:
                self.smtp_client = None


def main():
    """Main application entry point."""
    
    # Load configuration
    config = Config()
    if not config.validate():
        sys.exit(1)
    
    # Create SMTP OAuth client
    smtp_client = SMTPOAuthClient(config)
    
    try:
        # Authenticate using client credentials flow (OAuth2 authorization code flow)
        print("Authenticating using Client Credentials Flow (OAuth2 authorization code flow)...")
        if not smtp_client.authenticate_with_client_credentials():
            logger.error("Client credentials authentication failed")
            sys.exit(1)
        
        # Connect to SMTP server
        if not smtp_client.connect_smtp():
            logger.error("SMTP connection failed")
            sys.exit(1)
        
        # Test connection
        if not smtp_client.test_smtp_connection():
            logger.error("SMTP connection test failed")
            sys.exit(1)
        
        # Ask if user wants to send test email
        send_test = input("Send test email? (y/n): ").strip().lower()
        
        if send_test == 'y':
            recipient = input("Enter recipient email: ").strip()
            if recipient:
                success = smtp_client.send_test_email(
                    recipient,
                    "OAuth SMTP Test - Python",
                    "This is a test email sent using OAuth 2.0 Client Credentials flow with Microsoft Exchange from Python."
                )
                
                if not success:
                    logger.error("Failed to send test email")
                    sys.exit(1)
        
        print("OAuth SMTP test completed successfully!")
        
    except KeyboardInterrupt:
        print("\nOperation cancelled by user")
    except Exception as e:
        logger.error(f"Unexpected error: {str(e)}")
        sys.exit(1)
    finally:
        smtp_client.close()


if __name__ == "__main__":
    main()