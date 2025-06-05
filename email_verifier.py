import smtplib
import dns.resolver
import socket
import time
import re
from typing import Tuple


class EmailVerifier:
    """
    Email verification class using SMTP validation
    """
    
    def __init__(self, timeout: int = 10, sender_email: str = "test@example.com"):
        """
        Initialize email verifier
        
        Args:
            timeout: SMTP connection timeout in seconds
            sender_email: Email to use as sender for SMTP verification
        """
        self.timeout = timeout
        self.sender_email = sender_email
        
    def verify_email(self, email: str) -> Tuple[bool, str]:
        """
        Verify if an email address exists using SMTP validation
        
        Args:
            email: Email address to verify
            
        Returns:
            Tuple of (is_valid: bool, details: str)
        """
        try:
            # Basic format validation
            if not self._is_valid_email_format(email):
                return False, "Invalid email format"
            
            # Extract domain
            domain = email.split('@')[1].lower()
            
            # Get MX records
            try:
                mx_records = dns.resolver.resolve(domain, 'MX')
                mx_record = str(mx_records[0].exchange).rstrip('.')
            except (dns.resolver.NXDOMAIN, dns.resolver.NoAnswer, Exception) as e:
                return False, f"No MX record found for domain: {domain}"
            
            # SMTP verification
            try:
                # Create SMTP connection
                server = smtplib.SMTP(timeout=self.timeout)
                server.set_debuglevel(0)
                
                # Connect to MX server
                server.connect(mx_record, 25)
                
                # SMTP conversation
                server.helo(socket.gethostname())
                server.mail(self.sender_email)
                
                # The key part - check if recipient exists
                code, message = server.rcpt(email)
                server.quit()
                
                # Analyze response code
                if code == 250:
                    return True, "Email address exists"
                elif code == 550:
                    return False, "Email address does not exist"
                elif code == 552:
                    return False, "Mailbox full or quota exceeded"
                elif code == 553:
                    return False, "Invalid email address"
                else:
                    return False, f"SMTP error code: {code}"
                    
            except smtplib.SMTPConnectError:
                return False, "Cannot connect to SMTP server"
            except smtplib.SMTPServerDisconnected:
                return False, "SMTP server disconnected"
            except smtplib.SMTPRecipientsRefused:
                return False, "Recipient refused"
            except socket.timeout:
                return False, "SMTP connection timeout"
            except Exception as smtp_e:
                return False, f"SMTP error: {str(smtp_e)}"
                
        except Exception as e:
            return False, "error"
    
    def _is_valid_email_format(self, email: str) -> bool:
        """
        Validate email format using regex
        
        Args:
            email: Email address to validate
            
        Returns:
            bool: True if format is valid
        """
        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return re.match(pattern, email) is not None
    
    def verify_email_batch(self, emails: list, delay: float = 1.0) -> dict:
        """
        Verify multiple emails with rate limiting
        
        Args:
            emails: List of email addresses to verify
            delay: Delay between verifications in seconds
            
        Returns:
            dict: Email verification results
        """
        results = {}
        
        for email in emails:
            is_valid, details = self.verify_email(email)
            results[email] = {
                'is_valid': is_valid,
                'details': details
            }
            
            # Rate limiting
            if email != emails[-1]:  # Don't delay after last email
                time.sleep(delay)
        
        return results
    
    def get_domain_info(self, email: str) -> dict:
        """
        Get domain information for an email address
        
        Args:
            email: Email address
            
        Returns:
            dict: Domain information including MX records
        """
        try:
            domain = email.split('@')[1].lower()
            
            # Get MX records
            try:
                mx_records = dns.resolver.resolve(domain, 'MX')
                mx_list = [str(record.exchange).rstrip('.') for record in mx_records]
            except:
                mx_list = []
            
            # Get A records (domain resolution)
            try:
                a_records = dns.resolver.resolve(domain, 'A')
                a_list = [str(record) for record in a_records]
            except:
                a_list = []
            
            return {
                'domain': domain,
                'mx_records': mx_list,
                'a_records': a_list,
                'has_mx': len(mx_list) > 0,
                'resolvable': len(a_list) > 0
            }
            
        except Exception as e:
            return {
                'domain': 'unknown',
                'mx_records': [],
                'a_records': [],
                'has_mx': False,
                'resolvable': False,
                'error': str(e)
            }
