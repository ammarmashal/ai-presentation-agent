@echo off
echo Installing pyngrok...
pip install pyngrok

echo Starting Flask app with public URL...
python run_public.py
pause