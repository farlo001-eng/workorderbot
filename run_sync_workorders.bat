@echo off
cd /d I:\PycharmProjects\WorkOrderBot
echo [%date% %time%] Starting sync... >> sync_log.txt
py sync_workorders.py >> sync_log.txt 2>&1
echo [%date% %time%] Sync finished. >> sync_log.txt
