import json
import logging
import os
import uuid
from datetime import datetime
from typing import Dict, Any, List, Optional
from logging.handlers import RotatingFileHandler

# Import email detection utility
try:
    from email_utils import get_default_email_with_fallback
except ImportError:
    def get_default_email_with_fallback():
        """Fallback function if email_utils is not available"""
        try:
            import getpass
            username = getpass.getuser()
            return f"{username}@timesinternet.in"
        except:
            return "user@timesinternet.in"

class LineItemLogger:
    def __init__(self):
        self.logs_dir = "logs"
        os.makedirs(self.logs_dir, exist_ok=True)
        
        # Configure different loggers for different purposes
        self.user_logger = self._setup_logger('user_actions', 'user_actions.log')
        self.line_logger = self._setup_logger('line_creation', 'line_creation.log')
        self.creative_logger = self._setup_logger('creative_actions', 'creative_actions.log')
        self.error_logger = self._setup_logger('error_tracking', 'error_tracking.log')
        self.performance_logger = self._setup_logger('performance', 'performance.log')
        self.analytics_logger = self._setup_logger('analytics', 'analytics.json')
        
        # Generate session ID for this logging session
        self.session_id = str(uuid.uuid4())
        
        # Get system email ID
        self.system_email = get_default_email_with_fallback()
        
        # Add session divider to all log files
        self.add_session_divider()
        
    def _setup_logger(self, name: str, filename: str) -> logging.Logger:
        """Set up a logger with rotating file handler"""
        logger = logging.getLogger(name)
        logger.setLevel(logging.INFO)
        
        # Clear existing handlers
        logger.handlers.clear()
        
        # Create file handler with rotation
        file_path = os.path.join(self.logs_dir, filename)
        handler = RotatingFileHandler(
            file_path,
            maxBytes=10*1024*1024,  # 10MB
            backupCount=5
        )
        
        # Create formatter
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        handler.setFormatter(formatter)
        
        # Add handler to logger
        logger.addHandler(handler)
        
        # Prevent propagation to avoid duplicate logs
        logger.propagate = False
        
        return logger
    
    def add_session_divider(self):
        """Add a visual divider to all log files to separate new sessions"""
        timestamp = self.get_current_timestamp()
        divider_line = "=" * 80
        session_header = f"NEW SESSION STARTED - {timestamp}"
        session_details = f"Session ID: {self.session_id} | User: {self.system_email}"
        
        # Add divider to all loggers
        loggers = [
            self.user_logger,
            self.line_logger,
            self.creative_logger,
            self.error_logger,
            self.performance_logger
        ]
        
        for logger in loggers:
            logger.info("")  # Empty line
            logger.info(divider_line)
            logger.info(session_header)
            logger.info(session_details)
            logger.info(divider_line)
            logger.info("")  # Empty line
        
        # Add divider to analytics JSON
        divider_entry = {
            "timestamp": timestamp,
            "event_type": "SESSION_DIVIDER",
            "session_id": self.session_id,
            "system_email": self.system_email,
            "message": "New logging session started"
        }
        self.analytics_logger.info(json.dumps(divider_entry))
    
    def get_current_timestamp(self) -> str:
        """Get current timestamp in ISO format"""
        return datetime.now().isoformat()
    
    def log_user_input(self, user_data: Dict[str, Any], session_id: str = None):
        """
        Log user input and form data
        
        Args:
            user_data: Dictionary containing user input data
            session_id: Session identifier
        """
        timestamp = self.get_current_timestamp()
        
        log_entry = {
            "timestamp": timestamp,
            "event_type": "USER_INPUT",
            "session_id": session_id,
            "system_email": self.system_email,
            "user_data": user_data
        }
        
        self.user_logger.info(f"📋 USER INPUT CAPTURED")
        self.user_logger.info(f"User: {user_data.get('email', 'N/A')}")
        self.user_logger.info(f"Expresso ID: {user_data.get('expresso_id', 'N/A')}")
        self.user_logger.info(f"Order ID: {user_data.get('order_id', 'N/A')}")
        self.user_logger.info(f"Line Name: {user_data.get('line_name', 'N/A')}")
        self.user_logger.info(f"Site: {user_data.get('site', 'N/A')}")
        self.user_logger.info(f"Platform: {user_data.get('platform', 'N/A')}")
        self.user_logger.info(f"Geography: {user_data.get('geography', 'N/A')}")
        self.user_logger.info(f"Fcap: {user_data.get('fcap', 'N/A')}")
        self.user_logger.info(f"Currency: {user_data.get('currency', 'N/A')}")
        self.user_logger.info(f"Impressions: {user_data.get('impressions', 'N/A')}")
        self.user_logger.info(f"Destination URL: {user_data.get('destination_url', 'N/A')}")
        self.user_logger.info(f"Landing Page: {user_data.get('landing_page', 'N/A')}")
        self.user_logger.info(f"Impression Tracker: {user_data.get('impression_tracker', 'N/A')}")
        self.user_logger.info(f"Tracking Tag: {user_data.get('tracking_tag', 'N/A')}")
        self.user_logger.info(f"Banner Video: {user_data.get('banner_video', 'N/A')}")
        self.user_logger.info(f"Order Option: {user_data.get('order_option', 'N/A')}")
        self.user_logger.info(f"Uploaded Files: {user_data.get('uploaded_files', 'N/A')}")
        
        self.analytics_logger.info(json.dumps(log_entry))
    
    def log_line_creation_start(self, order_id: str, line_data: Dict[str, Any], line_name: str, session_id: str = None):
        """
        Log the start of line item creation
        
        Args:
            order_id: Order ID
            line_data: Line item data
            line_name: Line item name
            session_id: Session identifier
        """
        timestamp = self.get_current_timestamp()
        
        log_entry = {
            "timestamp": timestamp,
            "event_type": "LINE_CREATION_START",
            "session_id": session_id,
            "system_email": self.system_email,
            "order_id": order_id,
            "line_name": line_name,
            "line_data": line_data
        }
        
        self.line_logger.info(f"🚀 LINE ITEM CREATION STARTED")
        self.line_logger.info(f"Order ID: {order_id}")
        self.line_logger.info(f"Line Name: {line_name}")
        self.line_logger.info(f"Session ID: {session_id}")
        
        self.analytics_logger.info(json.dumps(log_entry))
    
    def log_line_creation_success(self, line_id: str, creative_ids: List[str], order_id: str, line_name: str, session_id: str = None):
        """
        Log successful line item creation
        
        Args:
            line_id: Created line item ID
            creative_ids: List of created creative IDs
            order_id: Order ID
            line_name: Line item name
            session_id: Session identifier
        """
        timestamp = self.get_current_timestamp()
        
        log_entry = {
            "timestamp": timestamp,
            "event_type": "LINE_CREATION_SUCCESS",
            "session_id": session_id,
            "system_email": self.system_email,
            "line_id": line_id,
            "creative_ids": creative_ids,
            "order_id": order_id,
            "line_name": line_name,
            "creative_count": len(creative_ids) if creative_ids else 0
        }
        
        self.line_logger.info(f"✅ LINE ITEM CREATED SUCCESSFULLY")
        self.line_logger.info(f"Line ID: {line_id}")
        self.line_logger.info(f"Order ID: {order_id}")
        self.line_logger.info(f"Line Name: {line_name}")
        self.line_logger.info(f"Creative Count: {len(creative_ids) if creative_ids else 0}")
        self.line_logger.info(f"Creative IDs: {creative_ids if creative_ids else 'None'}")
        
        self.analytics_logger.info(json.dumps(log_entry))
    
    def log_line_creation_error(self, error: Exception, line_name: str, order_id: str, session_id: str = None):
        """
        Log line item creation error
        
        Args:
            error: Exception that occurred
            line_name: Line item name
            order_id: Order ID
            session_id: Session identifier
        """
        timestamp = self.get_current_timestamp()
        
        log_entry = {
            "timestamp": timestamp,
            "event_type": "LINE_CREATION_ERROR",
            "session_id": session_id,
            "system_email": self.system_email,
            "error_message": str(error),
            "error_type": type(error).__name__,
            "line_name": line_name,
            "order_id": order_id
        }
        
        self.line_logger.error(f"❌ LINE ITEM CREATION FAILED")
        self.line_logger.error(f"Error: {str(error)}")
        self.line_logger.error(f"Line Name: {line_name}")
        self.line_logger.error(f"Order ID: {order_id}")
        
        self.error_logger.error(f"LINE CREATION ERROR: {str(error)}")
        self.error_logger.error(f"Line Name: {line_name}")
        self.error_logger.error(f"Order ID: {order_id}")
        
        self.analytics_logger.info(json.dumps(log_entry))
    
    def log_creative_creation(self, template_id: str, creative_id: str, size: str, asset_files: List[str] = None, session_id: str = None):
        """
        Log creative creation details
        
        Args:
            template_id: Template ID used for creative
            creative_id: Created creative ID
            size: Creative size
            asset_files: List of asset files used
            session_id: Session identifier
        """
        timestamp = self.get_current_timestamp()
        
        log_entry = {
            "timestamp": timestamp,
            "event_type": "CREATIVE_CREATION",
            "session_id": session_id,
            "system_email": self.system_email,
            "template_id": template_id,
            "creative_id": creative_id,
            "size": size,
            "asset_files": asset_files or []
        }
        
        self.creative_logger.info(f"🎨 CREATIVE CREATION")
        self.creative_logger.info(f"Size: {size}")
        self.creative_logger.info(f"Template ID: {template_id}")
        self.creative_logger.info(f"Creative ID: {creative_id}")
        self.creative_logger.info(f"Asset Files: {asset_files if asset_files else 'None'}")
        
        self.analytics_logger.info(json.dumps(log_entry))
    
    def log_placement_targeting(self, placement_data: Dict[str, Any], session_id: str = None):
        """
        Log placement targeting information
        
        Args:
            placement_data: Placement targeting data
            session_id: Session identifier
        """
        timestamp = self.get_current_timestamp()
        
        log_entry = {
            "timestamp": timestamp,
            "event_type": "PLACEMENT_TARGETING",
            "session_id": session_id,
            "system_email": self.system_email,
            "placement_data": placement_data
        }
        
        self.user_logger.info(f"🎯 PLACEMENT TARGETING")
        self.user_logger.info(f"Placement Count: {len(placement_data)}")
        
        for size, data in placement_data.items():
            # Handle different data types safely
            if isinstance(data, dict):
                placement_count = len(data.get('placement_ids', []))
                self.user_logger.info(f"Size {size}: {placement_count} placements")
            elif isinstance(data, list):
                self.user_logger.info(f"Size {size}: {len(data)} items")
            else:
                self.user_logger.info(f"Size {size}: {data}")
        
        self.analytics_logger.info(json.dumps(log_entry))
    
    def log_performance_metrics(self, metrics: Dict[str, Any], session_id: str = None):
        """
        Log performance metrics
        
        Args:
            metrics: Performance metrics data
            session_id: Session identifier
        """
        timestamp = self.get_current_timestamp()
        
        log_entry = {
            "timestamp": timestamp,
            "event_type": "PERFORMANCE_METRICS",
            "session_id": session_id,
            "system_email": self.system_email,
            "metrics": metrics
        }
        
        self.performance_logger.info(f"📊 PERFORMANCE METRICS")
        self.performance_logger.info(f"Session ID: {session_id}")
        self.performance_logger.info(f"Total Time: {metrics.get('total_time', 'N/A')}s")
        self.performance_logger.info(f"Data Processing: {metrics.get('data_processing_time', 'N/A')}s")
        self.performance_logger.info(f"Placement Lookup: {metrics.get('placement_lookup_time', 'N/A')}s")
        self.performance_logger.info(f"Line Creation: {metrics.get('line_creation_time', 'N/A')}s")
        self.performance_logger.info(f"Creative Creation: {metrics.get('creative_creation_time', 'N/A')}s")
        
        self.analytics_logger.info(json.dumps(log_entry))
    
    def log_geo_conflict(self, location_name: str, matches: List[Dict[str, Any]], session_id: str = None):
        """
        Log when multiple geo locations are found requiring CSM confirmation
        
        Args:
            location_name: The ambiguous location name
            matches: List of matching geo locations
            session_id: Session identifier
        """
        timestamp = self.get_current_timestamp()
        
        conflict_data = {
            'timestamp': timestamp,
            'session_id': session_id or self.session_id,
            'email': self.system_email,
            'event_type': 'geo_conflict',
            'location_name': location_name,
            'matches_count': len(matches),
            'matches': [
                {
                    'id': match['Id'],
                    'name': match['Name'],
                    'parent_region': match.get('ParentRegion', 'Unknown'),
                    'country_code': match['CountryCode'],
                    'type': match['Type']
                }
                for match in matches[:10]  # Limit to first 10 matches
            ],
            'requires_csm_confirmation': True,
            'status': 'pending_csm_confirmation'
        }
        
        # Log to user actions (since this affects user workflow)
        self.user_logger.warning(f"🚨 GEO CONFLICT DETECTED")
        self.user_logger.warning(f"Location: {location_name}")
        self.user_logger.warning(f"Multiple matches found: {len(matches)}")
        self.user_logger.warning(f"CSM confirmation required")
        
        # Log to error tracking for follow-up
        self.error_logger.warning(f"GEO_CONFLICT_REQUIRES_CSM | Location: {location_name} | Matches: {len(matches)}")
        
        # Add to analytics
        self.analytics_logger.info(json.dumps(conflict_data))

    def log_geo_auto_selection(self, location_name: str, selected_match: Dict[str, Any], parent_region: str, session_id: str = None):
        """
        Log when a geo location is automatically selected from multiple matches
        
        Args:
            location_name: The location name that had multiple matches
            selected_match: The geo location that was automatically selected
            parent_region: Parent region of the selected location
            session_id: Session identifier
        """
        timestamp = self.get_current_timestamp()
        
        selection_data = {
            'timestamp': timestamp,
            'session_id': session_id or self.session_id,
            'email': self.system_email,
            'event_type': 'geo_auto_selection',
            'location_name': location_name,
            'selected_geo': {
                'id': selected_match['Id'],
                'name': selected_match['Name'],
                'parent_region': parent_region,
                'country_code': selected_match['CountryCode'],
                'type': selected_match['Type']
            },
            'selection_method': 'first_match_with_warning'
        }
        
        # Log to user actions with prominent alert
        self.user_logger.warning(f"🚨 AUTOMATION ALERT: GEO AUTO-SELECTION")
        self.user_logger.warning(f"MULTIPLE LOCATIONS FOUND FOR: {location_name}")
        self.user_logger.warning(f"AUTOMATION SELECTED: {selected_match['Name']}, {parent_region}")
        self.user_logger.warning(f"SELECTED GEO ID: {selected_match['Id']}")
        self.user_logger.warning(f"REASON: First match from available options")
        self.user_logger.warning(f"ACTION REQUIRED: Review if different location needed")
        
        # Add to analytics
        self.analytics_logger.info(json.dumps(selection_data))

    def log_csm_confirmation(self, location_name: str, confirmed_geo_id: str, csm_email: str = None, session_id: str = None):
        """
        Log when CSM confirms a geo location selection
        
        Args:
            location_name: The location name that was confirmed
            confirmed_geo_id: The geo ID that was confirmed by CSM
            csm_email: Email of the CSM who confirmed
            session_id: Session identifier
        """
        timestamp = self.get_current_timestamp()
        
        confirmation_data = {
            'timestamp': timestamp,
            'session_id': session_id or self.session_id,
            'email': self.system_email,
            'event_type': 'csm_geo_confirmation',
            'location_name': location_name,
            'confirmed_geo_id': confirmed_geo_id,
            'csm_email': csm_email or 'csm@timesinternet.in',
            'resolution_status': 'confirmed'
        }
        
        # Log to user actions
        self.user_logger.info(f"✅ CSM GEO CONFIRMATION")
        self.user_logger.info(f"Location: {location_name}")
        self.user_logger.info(f"Confirmed Geo ID: {confirmed_geo_id}")
        self.user_logger.info(f"CSM: {csm_email or 'csm@timesinternet.in'}")
        
        # Add to analytics
        self.analytics_logger.info(json.dumps(confirmation_data))

# Global logger instance
logger = LineItemLogger() 