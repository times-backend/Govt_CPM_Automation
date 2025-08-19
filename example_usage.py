"""
Example usage of the comprehensive logging system for GAM line item creation.
This script demonstrates various logging scenarios and how to generate reports.
"""

import uuid
import time
from datetime import datetime
from logging_utils import logger
from log_monitor import LogMonitor

def example_user_input_logging():
    """
    Example of logging user input data
    """
    print("\n=== Example: User Input Logging ===")
    
    session_id = str(uuid.uuid4())
    
    # Sample user input data
    user_data = {
        'email': 'john.doe@example.com',
        'order_option': 'new',
        'order_id': None,
        'expresso_id': 271089,
        'label': 'Industry - FMCG',
        'line_name': 'TestCampaign_Brand_Awareness',
        'site': ['TOI', 'ETIMES'],
        'platforms': ['WEB', 'MWEB', 'AMP'],
        'geo': ['Mumbai', 'Delhi', 'Bangalore'],
        'fcap': '3',
        'currency': 'INR',
        'impressions': 1000000,
        'destination_url': 'https://example.com/campaign',
        'impression_tracker': 'https://tracker.example.com/imp',
        'tracking_tag': '<script>analytics.track("campaign");</script>',
        'banner_video': '',
        'uploaded_files': ['300x250.jpg', '728x90.jpg', '320x50.jpg']
    }
    
    # Log the user input
    logger.log_user_input(user_data, session_id)
    
    print(f"✅ User input logged with session ID: {session_id}")
    return session_id

def example_line_creation_logging():
    """
    Example of logging line creation process
    """
    print("\n=== Example: Line Creation Logging ===")
    
    session_id = str(uuid.uuid4())
    
    # Sample line item data
    line_item_data = {
        'CPM_Rate': 120.0,
        'impressions': 1000000,
        'Start_date': '2025-01-01 00:00:00',
        'End_date': '2025-01-31 23:59:59',
        'fcap': 0,
        'currency': 'INR',
        'site': ['TOI', 'ETIMES'],
        'platforms': ['WEB', 'MWEB'],
        'geoTargeting': ['Mumbai', 'Delhi'],
        'destination_url': 'https://example.com/campaign',
        'expresso_id': 271089
    }
    
    order_id = "3750012144"
    line_name = "TestCampaign_Brand_Awareness"
    
    # Log line creation start
    logger.log_line_creation_start(order_id, line_item_data, line_name, session_id)
    
    # Simulate line creation process
    time.sleep(1)
    
    # Log successful line creation
    line_item_id = "8967543210"
    creative_ids = ["1234567890", "1234567891", "1234567892"]
    
    logger.log_line_creation_success(line_item_id, creative_ids, line_name, session_id)
    
    # Log performance metrics
    logger.log_performance_metrics({
        'total_time': 2.5,
        'line_creation_time': 1.2,
        'creative_creation_time': 1.0,
        'placement_lookup_time': 0.3,
        'line_item_id': line_item_id,
        'creative_count': len(creative_ids)
    }, session_id)
    
    print(f"✅ Line creation logged with session ID: {session_id}")
    return session_id

def example_creative_logging():
    """
    Example of logging creative creation
    """
    print("\n=== Example: Creative Creation Logging ===")
    
    session_id = str(uuid.uuid4())
    
    # Sample creative creation scenarios
    creative_scenarios = [
        {
            'type': 'standard_banner',
            'size': '300x250',
            'template_id': 12399020,
            'creative_id': '1234567890',
            'asset_files': ['300x250.jpg']
        },
        {
            'type': 'rich_media',
            'size': '728x90',
            'template_id': 12344286,
            'creative_id': '1234567891',
            'asset_files': ['728x90.jpg', '728x90.js']
        },
        {
            'type': 'in_banner_video',
            'size': '300x250',
            'template_id': 12344286,
            'creative_id': '1234567892',
            'asset_files': ['300x250_video.mp4']
        }
    ]
    
    for creative_info in creative_scenarios:
        logger.log_creative_creation(creative_info, session_id)
    
    print(f"✅ Creative creation logged with session ID: {session_id}")
    return session_id

def example_error_logging():
    """
    Example of logging various error scenarios
    """
    print("\n=== Example: Error Logging ===")
    
    session_id = str(uuid.uuid4())
    
    # Simulate different error scenarios
    try:
        # Simulate a line creation error
        raise ValueError("Invalid geo targeting: 'Invalid City' not found")
    except Exception as e:
        logger.log_line_creation_error(e, "TestCampaign_Error", "3750012144", session_id)
    
    try:
        # Simulate a creative creation error
        raise FileNotFoundError("Creative file not found: 300x250.jpg")
    except Exception as e:
        logger.log_creative_error(e, {
            'type': 'standard_banner',
            'size': '300x250',
            'template_id': 12399020
        }, session_id)
    
    # Log system events
    logger.log_system_event('LOCATION_ERROR', {
        'location_name': 'Invalid City',
        'error_message': 'Location not found in geo targeting database',
        'suggested_locations': ['Mumbai', 'Delhi', 'Bangalore']
    }, session_id)
    
    print(f"✅ Error scenarios logged with session ID: {session_id}")
    return session_id

def example_cpd_logging():
    """
    Example of logging CPD multiple line creation
    """
    print("\n=== Example: CPD Multiple Lines Logging ===")
    
    session_id = str(uuid.uuid4())
    
    # Log CPD creation
    logger.log_cpd_multiple_lines({
        'base_line_name': 'TestCampaign_CPD',
        'total_lines': 30,
        'days_per_line': 1,
        'total_impressions': 3000000,
        'impressions_per_line': 100000,
        'order_id': '3750012144',
        'start_date': '2025-01-01',
        'end_date': '2025-01-30'
    }, session_id)
    
    # Log session summary
    logger.log_session_summary({
        'lines_created': 30,
        'creatives_created': 90,
        'errors': 2,
        'success_rate': 93.33,
        'total_time': 45.2
    }, session_id)
    
    print(f"✅ CPD logging completed with session ID: {session_id}")
    return session_id

def example_placement_logging():
    """
    Example of logging placement targeting
    """
    print("\n=== Example: Placement Targeting Logging ===")
    
    session_id = str(uuid.uuid4())
    
    # Log placement targeting
    logger.log_placement_targeting({
        'site_filter': ['TOI', 'ETIMES'],
        'platform_filter': ['WEB', 'MWEB', 'AMP'],
        'placement_count': 45,
        'size_groups': {
            '300x250': {'count': 15, 'sites': ['TOI', 'ETIMES']},
            '728x90': {'count': 20, 'sites': ['TOI', 'ETIMES']},
            '320x50': {'count': 10, 'sites': ['TOI', 'ETIMES']}
        },
        'contains_toi': True,
        'contains_et': False
    }, session_id)
    
    print(f"✅ Placement targeting logged with session ID: {session_id}")
    return session_id

def generate_sample_reports():
    """
    Generate sample reports using the log monitor
    """
    print("\n=== Generating Sample Reports ===")
    
    monitor = LogMonitor()
    
    # Generate daily report
    print("\n📊 Daily Report:")
    monitor.print_daily_summary()
    
    # Generate error report
    print("\n⚠️ Error Report:")
    try:
        error_report = monitor.generate_error_report(days=7)
        print(f"Total Errors (7 days): {error_report['total_errors']}")
        
        if error_report['most_common_errors']:
            print("Most Common Errors:")
            for error_type, count in error_report['most_common_errors'][:5]:
                print(f"  • {error_type}: {count}")
        else:
            print("No errors found in the last 7 days")
    except Exception as e:
        print(f"Error generating report: {e}")
    
    # Generate performance report
    print("\n⚡ Performance Report:")
    try:
        perf_report = monitor.generate_performance_report(days=30)
        print(f"Total Measurements (30 days): {perf_report['total_measurements']}")
        print(f"Avg Processing Time: {perf_report['avg_total_time']:.2f}s")
        print(f"Avg Line Creation Time: {perf_report['avg_line_creation_time']:.2f}s")
        print(f"Avg Creative Creation Time: {perf_report['avg_creative_creation_time']:.2f}s")
    except Exception as e:
        print(f"Error generating performance report: {e}")

def main():
    """
    Run all example scenarios
    """
    print("🚀 GAM Logging System Examples")
    print("=" * 50)
    
    # Run all examples
    session_ids = []
    
    session_ids.append(example_user_input_logging())
    session_ids.append(example_line_creation_logging())
    session_ids.append(example_creative_logging())
    session_ids.append(example_error_logging())
    session_ids.append(example_cpd_logging())
    session_ids.append(example_placement_logging())
    
    print(f"\n✅ All examples completed!")
    print(f"📋 Generated {len(session_ids)} example sessions")
    print(f"📁 Check the 'logs' directory for generated log files")
    
    # Generate sample reports
    generate_sample_reports()
    
    print("\n💡 Usage Tips:")
    print("1. Use 'python log_monitor.py --report daily' for daily reports")
    print("2. Use 'python log_monitor.py --report error --days 7' for error analysis")
    print("3. Use 'python log_monitor.py --report performance --days 30' for performance analysis")
    print("4. Use 'python log_monitor.py --report user --user email@example.com' for user activity")
    print("5. Add '--export' flag to save reports to Excel files")

if __name__ == "__main__":
    main() 