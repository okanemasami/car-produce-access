@echo off
REM カーセンサー問い合わせ漏れ防止用CSVをGoogle Driveにアップロードするバッチファイル
REM Windowsタスクスケジューラから実行します

REM スクリプトのディレクトリに移動
cd /d %~dp0

REM Pythonスクリプトを実行してログに記録
echo ========================================== >> upload_carsensor_inquiry_log.txt
echo 実行開始: %date% %time% >> upload_carsensor_inquiry_log.txt
echo ========================================== >> upload_carsensor_inquiry_log.txt

python upload_carsensor_inquiry.py >> upload_carsensor_inquiry_log.txt 2>&1

echo. >> upload_carsensor_inquiry_log.txt
echo 実行完了: %date% %time% >> upload_carsensor_inquiry_log.txt
echo. >> upload_carsensor_inquiry_log.txt
