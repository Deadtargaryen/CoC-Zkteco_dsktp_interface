from zk import ZK, const
import pandas as pd
from dateutil import parser
from datetime import datetime, time, timezone, timedelta
import threading
import time as time_module
import requests
import os


LAGOS_TZ = timezone(timedelta(hours=1))

class ZKTecoAttendance:
    def __init__(self, ip_address, port=4370, timeout=5, password=0,
                 api_url="https://coc4towns-attendance.vercel.app/api/attendance/device",
                 api_key="super-secret-key-here", poll_interval=60):
        self.ip_address = ip_address
        self.port = port
        self.timeout = timeout
        self.password = password
        self.zk = ZK(self.ip_address, port=self.port, timeout=self.timeout, password=self.password)
        self.conn = None
        self.users = {}

        # Sync-related
        self.api_url = api_url
        self.api_key = api_key
        self.poll_interval = poll_interval
        self.last_sync_file = "last_sync.txt"
        self.last_sync_time = None
        self.last_api_status = None
        self.sync_thread = None
        self.sync_running = False

    # ------------------ CONNECTION ------------------
    def connect(self):
        try:
            self.conn = self.zk.connect()
            self.load_users()
            print(f"Successfully connected to device at {self.ip_address}")

            # Start sync thread
            self.sync_running = True
            self.sync_thread = threading.Thread(target=self._sync_loop, daemon=True)
            self.sync_thread.start()
        except Exception as e:
            print(f"Error connecting to device: {str(e)}")
            self.conn = None

    def disconnect(self):
        self.sync_running = False
        if self.conn:
            self.conn.disconnect()
            print("Disconnected from device")
            self.conn = None
            self.users = {}

    # ------------------ USERS ------------------
    def load_users(self):
        if not self.conn:
            return
        try:
            users = self.conn.get_users()
            self.users = {user.user_id: user.name for user in users}
            print(f"Loaded {len(self.users)} users from device")
        except Exception as e:
            print(f"Error loading users: {str(e)}")

    # ------------------ ATTENDANCE ------------------
    def get_attendance_status(self, punch):
        punch_map = {0: "Check In", 1: "Check Out"}
        return punch_map.get(punch, f"Unknown Punch ({punch})")

    def get_attendance(self, start_date=None, end_date=None):
        if not self.conn:
            print("Not connected to device. Please connect first.")
            return None
        try:
            attendance = self.conn.get_attendance()
            if not attendance:
                print("No attendance records found")
                return None

            print(f"Retrieved {len(attendance)} attendance records")

            if start_date and not isinstance(start_date, datetime):
                start_date = datetime.combine(start_date, time.min)
            if end_date and not isinstance(end_date, datetime):
                end_date = datetime.combine(end_date, time.max)

            raw_records = []
            for att in attendance:
                dt = att.timestamp
                if start_date and end_date:
                    if not (start_date <= dt <= end_date):
                        continue
                user_name = self.users.get(str(att.user_id), "Unknown")
                raw_records.append({
                    'user_id': att.user_id,
                    'user_name': user_name,
                    'timestamp': dt,
                    'raw_status': att.status,
                    'punch': att.punch,
                    'status': self.get_attendance_status(att.punch)
                })

            df = pd.DataFrame(raw_records)
            if df.empty:
                return None

            df = df.sort_values(['user_id', 'timestamp'])

            grouped_records = []
            current_user = None
            current_date = None
            check_in = None
            check_out = None

            for _, row in df.iterrows():
                user_id = row['user_id']
                user_name = row['user_name']
                timestamp = row['timestamp']
                date = timestamp.date()
                punch = row['punch']

                if current_user != user_id or current_date != date:
                    if current_user is not None and check_in is not None:
                        grouped_records.append({
                            'user_id': current_user,
                            'user_name': current_user_name,
                            'date': current_date,
                            'check_in': check_in,
                            'check_out': check_out
                        })
                    current_user = user_id
                    current_user_name = user_name
                    current_date = date
                    check_in = None
                    check_out = None

                if punch == 0:
                    check_in = timestamp
                elif punch == 1:
                    check_out = timestamp

            if current_user is not None and check_in is not None:
                grouped_records.append({
                    'user_id': current_user,
                    'user_name': current_user_name,
                    'date': current_date,
                    'check_in': check_in,
                    'check_out': check_out
                })

            result_df = pd.DataFrame(grouped_records)

            if not result_df.empty:
                result_df['duration'] = result_df.apply(
                    lambda row: (row['check_out'] - row['check_in']).total_seconds() / 3600 
                    if pd.notnull(row['check_out']) else None,
                    axis=1
                )

            print(f"\nGrouped into {len(result_df)} attendance records")
            if not result_df.empty:
                print("\nSample of grouped records:")
                print(result_df.head())
            return result_df

        except Exception as e:
            print(f"Error retrieving attendance records: {str(e)}")
            return None

    # ------------------ SYNC LOGIC ------------------
    def _load_last_sync(self):
        try:
            if os.path.exists(self.last_sync_file):
                with open(self.last_sync_file, "r") as f:
                    return datetime.fromisoformat(f.read().strip())
        except Exception:
            return None
        return None

    def _save_last_sync(self, ts):
        try:
            with open(self.last_sync_file, "w") as f:
                f.write(ts.isoformat())
        except Exception as e:
            print(f"Error saving last sync time: {e}")

    def _send_log(self, user_id, timestamp):
        if timestamp.tzinfo is None:
            timestamp = timestamp.replace(tzinfo=LAGOS_TZ)
        payload = {"user_id": str(user_id), "timestamp": timestamp.isoformat()}
        headers = {"x-api-key": self.api_key, "Content-Type": "application/json"}
        try:
            r = requests.post(self.api_url, json=payload, headers=headers, timeout=50)
            self.last_api_status = f"{datetime.now()} → Status {r.status_code}"
            print(f"Sent log {payload} → Status {r.status_code}")
        except Exception as e:
            self.last_api_status = f"Error: {e}"
            print(f"Error sending log: {e}")

    def _sync_loop(self):
        print("Starting background sync loop...")
        last_sync = self._load_last_sync()
        new_last_sync = last_sync

        # Allowed days: Sunday(6), Monday(0), Wednesday(2), Friday(4)
        ALLOWED_DAYS = {6, 0, 2, 4}

        while self.sync_running:
            try:
                if not self.conn:
                    print("Not connected, skipping sync cycle...")
                    time_module.sleep(self.poll_interval)
                    continue

                logs = self.conn.get_attendance()
                logs.sort(key=lambda x: x.timestamp)

                for log in logs:
                    log_day = log.timestamp.weekday()
                    if log_day in ALLOWED_DAYS:
                        if last_sync is None or log.timestamp > last_sync:
                            self._send_log(log.user_id, log.timestamp)
                            if new_last_sync is None or log.timestamp > new_last_sync:
                                new_last_sync = log.timestamp
                    else:
                        print(f"Skipping log for {log.timestamp.strftime('%A')} ({log.timestamp})")

                if new_last_sync:
                    self._save_last_sync(new_last_sync)
                    self.last_sync_time = new_last_sync
                    print(f"Updated last sync time → {new_last_sync}")

            except Exception as e:
                print(f"Sync error: {e}")

            time_module.sleep(self.poll_interval)

    def get_sync_status(self):
        return {
            "last_sync_time": self.last_sync_time,
            "last_api_status": self.last_api_status
        }

# ------------------ MAIN TEST ------------------
def main():
    device_ip = "192.168.1.201"  # Replace with your device's IP address
    attendance_system = ZKTecoAttendance(device_ip)
    try:
        attendance_system.connect()
        if attendance_system.conn:
            attendance_records = attendance_system.get_attendance()
            if attendance_records is not None:
                print("\nAttendance Records:")
                print(attendance_records)
                attendance_records.to_csv('attendance_records.csv', index=False)
                print("\nRecords saved to 'attendance_records.csv'")

            print("Running... Press Ctrl+C to exit.")
            while True:
                time_module.sleep(10)
    except KeyboardInterrupt:
        print("Exiting...")
    finally:
        attendance_system.disconnect()

if __name__ == "__main__":
    main()
