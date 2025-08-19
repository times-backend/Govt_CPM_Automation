import json
import os
import pandas as pd
from datetime import datetime, timedelta
import argparse
from typing import Dict, List, Any, Optional
import matplotlib.pyplot as plt
import seaborn as sns
from logging_utils import logger

class LogMonitor:
    """
    Log monitoring and reporting utility for GAM line item creation
    Provides various analysis and reporting capabilities
    """
    
    def __init__(self, log_directory: str = "logs"):
        """
        Initialize the log monitor
        
        Args:
            log_directory: Directory containing log files
        """
        self.log_directory = log_directory
        self.analytics_file = os.path.join(log_directory, "analytics.json")
        
    def load_analytics_data(self, start_date: Optional[str] = None, 
                           end_date: Optional[str] = None) -> List[Dict]:
        """
        Load analytics data from JSON log file
        
        Args:
            start_date: Start date in ISO format (optional)
            end_date: End date in ISO format (optional)
            
        Returns:
            List of log entries
        """
        if not os.path.exists(self.analytics_file):
            print(f"Analytics file not found: {self.analytics_file}")
            return []
        
        entries = []
        try:
            with open(self.analytics_file, 'r') as f:
                for line in f:
                    try:
                        entry = json.loads(line.strip())
                        
                        # Filter by date range if specified
                        if start_date or end_date:
                            entry_date = entry.get('timestamp', '')
                            if start_date and entry_date < start_date:
                                continue
                            if end_date and entry_date > end_date:
                                continue
                        
                        entries.append(entry)
                    except json.JSONDecodeError:
                        continue
        except Exception as e:
            print(f"Error reading analytics file: {e}")
            return []
        
        return entries
    
    def generate_daily_report(self, date: str = None) -> Dict[str, Any]:
        """
        Generate daily activity report
        
        Args:
            date: Date in YYYY-MM-DD format (defaults to today)
            
        Returns:
            Dictionary with daily statistics
        """
        if not date:
            date = datetime.now().strftime("%Y-%m-%d")
        
        start_date = f"{date}T00:00:00"
        end_date = f"{date}T23:59:59"
        
        entries = self.load_analytics_data(start_date, end_date)
        
        report = {
            "date": date,
            "total_events": len(entries),
            "user_inputs": 0,
            "line_creations": 0,
            "line_successes": 0,
            "line_errors": 0,
            "creative_creations": 0,
            "creative_errors": 0,
            "cpd_creations": 0,
            "performance_metrics": [],
            "error_summary": {},
            "user_activity": {},
            "hourly_distribution": {}
        }
        
        for entry in entries:
            event_type = entry.get("event_type", "")
            timestamp = entry.get("timestamp", "")
            
            # Count by event type
            if event_type == "USER_INPUT":
                report["user_inputs"] += 1
                user = entry.get("user_data", {}).get("email", "unknown")
                report["user_activity"][user] = report["user_activity"].get(user, 0) + 1
                
            elif event_type == "LINE_CREATION_START":
                report["line_creations"] += 1
                
            elif event_type == "LINE_CREATION_SUCCESS":
                report["line_successes"] += 1
                
            elif event_type == "LINE_CREATION_ERROR":
                report["line_errors"] += 1
                error_type = entry.get("error_type", "Unknown")
                report["error_summary"][error_type] = report["error_summary"].get(error_type, 0) + 1
                
            elif event_type == "CREATIVE_CREATION":
                report["creative_creations"] += 1
                
            elif event_type == "CREATIVE_ERROR":
                report["creative_errors"] += 1
                error_type = entry.get("error_type", "Unknown")
                report["error_summary"][error_type] = report["error_summary"].get(error_type, 0) + 1
                
            elif event_type == "CPD_MULTIPLE_LINES":
                report["cpd_creations"] += 1
                
            elif event_type == "PERFORMANCE_METRICS":
                report["performance_metrics"].append(entry.get("metrics", {}))
            
            # Hourly distribution
            if timestamp:
                try:
                    hour = datetime.fromisoformat(timestamp).hour
                    report["hourly_distribution"][hour] = report["hourly_distribution"].get(hour, 0) + 1
                except:
                    pass
        
        return report
    
    def generate_user_report(self, user_email: str, days: int = 30) -> Dict[str, Any]:
        """
        Generate user activity report
        
        Args:
            user_email: Email of the user
            days: Number of days to look back
            
        Returns:
            Dictionary with user statistics
        """
        end_date = datetime.now()
        start_date = end_date - timedelta(days=days)
        
        entries = self.load_analytics_data(
            start_date.isoformat(), 
            end_date.isoformat()
        )
        
        user_entries = [
            entry for entry in entries 
            if entry.get("user_data", {}).get("email") == user_email or
               entry.get("details", {}).get("user") == user_email
        ]
        
        report = {
            "user_email": user_email,
            "period_days": days,
            "total_submissions": 0,
            "successful_lines": 0,
            "failed_lines": 0,
            "total_creatives": 0,
            "avg_processing_time": 0,
            "most_used_sites": {},
            "most_used_platforms": {},
            "expresso_ids": [],
            "line_names": [],
            "activity_by_date": {}
        }
        
        processing_times = []
        
        for entry in user_entries:
            event_type = entry.get("event_type", "")
            timestamp = entry.get("timestamp", "")
            
            if event_type == "USER_INPUT":
                report["total_submissions"] += 1
                user_data = entry.get("user_data", {})
                
                # Track sites and platforms
                sites = user_data.get("site", [])
                platforms = user_data.get("platforms", [])
                
                for site in sites:
                    report["most_used_sites"][site] = report["most_used_sites"].get(site, 0) + 1
                
                for platform in platforms:
                    report["most_used_platforms"][platform] = report["most_used_platforms"].get(platform, 0) + 1
                
                # Track expresso IDs and line names
                expresso_id = user_data.get("expresso_id")
                line_name = user_data.get("line_name")
                
                if expresso_id:
                    report["expresso_ids"].append(expresso_id)
                if line_name:
                    report["line_names"].append(line_name)
                
                # Activity by date
                if timestamp:
                    try:
                        date = datetime.fromisoformat(timestamp).date().isoformat()
                        report["activity_by_date"][date] = report["activity_by_date"].get(date, 0) + 1
                    except:
                        pass
            
            elif event_type == "LINE_CREATION_SUCCESS":
                report["successful_lines"] += 1
                creative_count = entry.get("creative_count", 0)
                report["total_creatives"] += creative_count
                
            elif event_type == "LINE_CREATION_ERROR":
                report["failed_lines"] += 1
                
            elif event_type == "PERFORMANCE_METRICS":
                metrics = entry.get("metrics", {})
                total_time = metrics.get("total_time", 0)
                if total_time:
                    processing_times.append(total_time)
        
        # Calculate average processing time
        if processing_times:
            report["avg_processing_time"] = sum(processing_times) / len(processing_times)
        
        return report
    
    def generate_error_report(self, days: int = 7) -> Dict[str, Any]:
        """
        Generate error analysis report
        
        Args:
            days: Number of days to analyze
            
        Returns:
            Dictionary with error statistics
        """
        end_date = datetime.now()
        start_date = end_date - timedelta(days=days)
        
        entries = self.load_analytics_data(
            start_date.isoformat(), 
            end_date.isoformat()
        )
        
        error_entries = [
            entry for entry in entries 
            if entry.get("event_type", "").endswith("_ERROR")
        ]
        
        report = {
            "period_days": days,
            "total_errors": len(error_entries),
            "error_types": {},
            "error_frequency_by_hour": {},
            "error_patterns": {},
            "most_common_errors": [],
            "error_timeline": []
        }
        
        for entry in error_entries:
            error_type = entry.get("error_type", "Unknown")
            timestamp = entry.get("timestamp", "")
            event_type = entry.get("event_type", "")
            
            # Count by error type
            report["error_types"][error_type] = report["error_types"].get(error_type, 0) + 1
            
            # Error frequency by hour
            if timestamp:
                try:
                    hour = datetime.fromisoformat(timestamp).hour
                    report["error_frequency_by_hour"][hour] = report["error_frequency_by_hour"].get(hour, 0) + 1
                except:
                    pass
            
            # Error patterns
            if event_type not in report["error_patterns"]:
                report["error_patterns"][event_type] = {}
            
            report["error_patterns"][event_type][error_type] = report["error_patterns"][event_type].get(error_type, 0) + 1
            
            # Timeline
            report["error_timeline"].append({
                "timestamp": timestamp,
                "event_type": event_type,
                "error_type": error_type,
                "error_message": entry.get("error_message", "")
            })
        
        # Sort most common errors
        report["most_common_errors"] = sorted(
            report["error_types"].items(), 
            key=lambda x: x[1], 
            reverse=True
        )[:10]
        
        return report
    
    def generate_performance_report(self, days: int = 30) -> Dict[str, Any]:
        """
        Generate performance analysis report
        
        Args:
            days: Number of days to analyze
            
        Returns:
            Dictionary with performance statistics
        """
        end_date = datetime.now()
        start_date = end_date - timedelta(days=days)
        
        entries = self.load_analytics_data(
            start_date.isoformat(), 
            end_date.isoformat()
        )
        
        perf_entries = [
            entry for entry in entries 
            if entry.get("event_type") == "PERFORMANCE_METRICS"
        ]
        
        report = {
            "period_days": days,
            "total_measurements": len(perf_entries),
            "avg_total_time": 0,
            "avg_line_creation_time": 0,
            "avg_creative_creation_time": 0,
            "avg_placement_lookup_time": 0,
            "performance_trends": [],
            "slowest_operations": [],
            "fastest_operations": []
        }
        
        if not perf_entries:
            return report
        
        total_times = []
        line_times = []
        creative_times = []
        placement_times = []
        
        for entry in perf_entries:
            metrics = entry.get("metrics", {})
            
            total_time = metrics.get("total_time", 0)
            line_time = metrics.get("line_creation_time", 0)
            creative_time = metrics.get("creative_creation_time", 0)
            placement_time = metrics.get("placement_lookup_time", 0)
            
            if total_time:
                total_times.append(total_time)
                report["performance_trends"].append({
                    "timestamp": entry.get("timestamp"),
                    "total_time": total_time,
                    "session_id": entry.get("session_id")
                })
            
            if line_time:
                line_times.append(line_time)
            if creative_time:
                creative_times.append(creative_time)
            if placement_time:
                placement_times.append(placement_time)
        
        # Calculate averages
        if total_times:
            report["avg_total_time"] = sum(total_times) / len(total_times)
        if line_times:
            report["avg_line_creation_time"] = sum(line_times) / len(line_times)
        if creative_times:
            report["avg_creative_creation_time"] = sum(creative_times) / len(creative_times)
        if placement_times:
            report["avg_placement_lookup_time"] = sum(placement_times) / len(placement_times)
        
        # Find slowest and fastest operations
        sorted_times = sorted(total_times, reverse=True)
        if len(sorted_times) >= 5:
            report["slowest_operations"] = sorted_times[:5]
            report["fastest_operations"] = sorted_times[-5:]
        
        return report
    
    def export_to_excel(self, report_type: str = "daily", **kwargs) -> str:
        """
        Export report to Excel file
        
        Args:
            report_type: Type of report ("daily", "user", "error", "performance")
            **kwargs: Additional arguments for report generation
            
        Returns:
            Path to the generated Excel file
        """
        if report_type == "daily":
            report = self.generate_daily_report(kwargs.get("date"))
        elif report_type == "user":
            report = self.generate_user_report(kwargs.get("user_email"), kwargs.get("days", 30))
        elif report_type == "error":
            report = self.generate_error_report(kwargs.get("days", 7))
        elif report_type == "performance":
            report = self.generate_performance_report(kwargs.get("days", 30))
        else:
            raise ValueError(f"Unknown report type: {report_type}")
        
        # Create DataFrame from report
        df = pd.DataFrame([report])
        
        # Generate filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"gam_log_report_{report_type}_{timestamp}.xlsx"
        
        # Save to Excel
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Summary', index=False)
            
            # Add detailed sheets for complex data
            if report_type == "daily" and report.get("performance_metrics"):
                perf_df = pd.DataFrame(report["performance_metrics"])
                perf_df.to_excel(writer, sheet_name='Performance', index=False)
            
            if report_type == "error" and report.get("error_timeline"):
                error_df = pd.DataFrame(report["error_timeline"])
                error_df.to_excel(writer, sheet_name='Error Timeline', index=False)
        
        print(f"Report exported to: {filename}")
        return filename
    
    def print_daily_summary(self, date: str = None):
        """
        Print a formatted daily summary to console
        
        Args:
            date: Date in YYYY-MM-DD format (defaults to today)
        """
        report = self.generate_daily_report(date)
        
        print("\n" + "="*60)
        print(f"📊 DAILY ACTIVITY REPORT - {report['date']}")
        print("="*60)
        
        print(f"📈 Total Events: {report['total_events']}")
        print(f"👥 User Inputs: {report['user_inputs']}")
        print(f"🚀 Line Creations: {report['line_creations']}")
        print(f"✅ Line Successes: {report['line_successes']}")
        print(f"❌ Line Errors: {report['line_errors']}")
        print(f"🎨 Creative Creations: {report['creative_creations']}")
        print(f"🔥 Creative Errors: {report['creative_errors']}")
        print(f"📅 CPD Creations: {report['cpd_creations']}")
        
        if report['user_activity']:
            print(f"\n👤 User Activity:")
            for user, count in sorted(report['user_activity'].items(), key=lambda x: x[1], reverse=True):
                print(f"  • {user}: {count} submissions")
        
        if report['error_summary']:
            print(f"\n⚠️ Error Summary:")
            for error_type, count in sorted(report['error_summary'].items(), key=lambda x: x[1], reverse=True):
                print(f"  • {error_type}: {count} occurrences")
        
        if report['performance_metrics']:
            avg_time = sum(m.get('total_time', 0) for m in report['performance_metrics']) / len(report['performance_metrics'])
            print(f"\n⚡ Performance:")
            print(f"  • Average Processing Time: {avg_time:.2f} seconds")
            print(f"  • Performance Measurements: {len(report['performance_metrics'])}")
        
        print("="*60)

def main():
    """
    Command-line interface for log monitoring
    """
    parser = argparse.ArgumentParser(description='GAM Log Monitor')
    parser.add_argument('--report', choices=['daily', 'user', 'error', 'performance'], 
                       default='daily', help='Type of report to generate')
    parser.add_argument('--date', help='Date for daily report (YYYY-MM-DD)')
    parser.add_argument('--user', help='User email for user report')
    parser.add_argument('--days', type=int, default=30, help='Number of days to analyze')
    parser.add_argument('--export', action='store_true', help='Export report to Excel')
    
    args = parser.parse_args()
    
    monitor = LogMonitor()
    
    if args.report == 'daily':
        if args.export:
            monitor.export_to_excel('daily', date=args.date)
        else:
            monitor.print_daily_summary(args.date)
    
    elif args.report == 'user':
        if not args.user:
            print("Error: --user email is required for user report")
            return
        
        if args.export:
            monitor.export_to_excel('user', user_email=args.user, days=args.days)
        else:
            report = monitor.generate_user_report(args.user, args.days)
            print(f"\n📊 User Report for {args.user} ({args.days} days):")
            print(f"Total Submissions: {report['total_submissions']}")
            print(f"Successful Lines: {report['successful_lines']}")
            print(f"Failed Lines: {report['failed_lines']}")
            print(f"Total Creatives: {report['total_creatives']}")
            print(f"Avg Processing Time: {report['avg_processing_time']:.2f}s")
    
    elif args.report == 'error':
        if args.export:
            monitor.export_to_excel('error', days=args.days)
        else:
            report = monitor.generate_error_report(args.days)
            print(f"\n⚠️ Error Report ({args.days} days):")
            print(f"Total Errors: {report['total_errors']}")
            print("Most Common Errors:")
            for error_type, count in report['most_common_errors']:
                print(f"  • {error_type}: {count}")
    
    elif args.report == 'performance':
        if args.export:
            monitor.export_to_excel('performance', days=args.days)
        else:
            report = monitor.generate_performance_report(args.days)
            print(f"\n⚡ Performance Report ({args.days} days):")
            print(f"Total Measurements: {report['total_measurements']}")
            print(f"Avg Total Time: {report['avg_total_time']:.2f}s")
            print(f"Avg Line Creation Time: {report['avg_line_creation_time']:.2f}s")
            print(f"Avg Creative Creation Time: {report['avg_creative_creation_time']:.2f}s")

if __name__ == "__main__":
    main() 