set -e

python3 /app/updatedpy.py

jupyter nbconvert \
  --to notebook \
  --execute notebooks/CSCI316_Project_Modeling.ipynb \
  --output executed_modeling.ipynb \
  --output-dir /app/out