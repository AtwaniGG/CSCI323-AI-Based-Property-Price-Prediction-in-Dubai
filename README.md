1) Prerequisites
- Install Docker Desktop for Windows
- Open Docker Desktop and make sure it is running
- Open a terminal in the project folder 
Your folder should include:
- Dockerfile
- requirements.txt
- run_pipeline.sh
- notebooks\CSCI316_Project_Modeling.ipynb
- data\ (contains the cleaned CSV used by the notebook)

Build
docker build --platform=linux/amd64 -t csci316-spark-ml .

Run: (FAST demo)
mkdir out -Force | Out-Null
docker run --rm -it --platform=linux/amd64 `
  -e FAST_DOCKER=1 `
  -v "${PWD}\out:/app/out" `
  csci316-spark-ml

Run: (4gb RAM)
docker run --rm -it --platform linux/amd64 `
  -e FAST_DOCKER=1 `
  -v "${PWD}:/app" `
  --memory="4g" `
  --memory-swap="8g" `
  csci316-spark-ml `
  python3 PricePulseAI.py
  
Check output
dir out

Expected: executed_modeling.ipynb