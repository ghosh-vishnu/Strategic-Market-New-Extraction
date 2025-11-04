#!/usr/bin/env python
"""
Custom Django server runner that handles broken pipe errors gracefully.
"""

import os
import sys
import signal
import logging
from django.core.management import execute_from_command_line

def signal_handler(signum, frame):
    """Handle system signals gracefully."""
    if signum == signal.SIGPIPE:
        # Ignore broken pipe signals
        pass
    else:
        sys.exit(0)

def main():
    """Main function to run the Django server with broken pipe handling."""
    
    # Set up signal handlers
    signal.signal(signal.SIGPIPE, signal.SIG_DFL)  # Default handling for broken pipes
    signal.signal(signal.SIGINT, signal_handler)   # Handle Ctrl+C gracefully
    signal.signal(signal.SIGTERM, signal_handler)  # Handle termination gracefully
    
    # Configure logging to reduce broken pipe noise
    logging.basicConfig(
        level=logging.WARNING,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    
    # Suppress broken pipe warnings
    logging.getLogger('django.server').setLevel(logging.ERROR)
    
    # Set environment variables
    os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'excel_backend.settings')
    
    # Run Django management command
    try:
        execute_from_command_line(sys.argv)
    except (BrokenPipeError, ConnectionResetError, OSError) as e:
        if 'Broken pipe' in str(e) or 'Connection reset' in str(e):
            # Silently handle broken pipe errors
            pass
        else:
            raise

if __name__ == '__main__':
    main()
