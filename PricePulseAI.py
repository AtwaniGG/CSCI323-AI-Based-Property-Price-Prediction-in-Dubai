# Commands to run in PowerShell inside project folder for Docker
# docker build -f Dockerfile -t csci316-project .
'''
docker run --rm -it --platform linux/amd64 `
   -e FAST_DOCKER=1 `
   -v "${PWD}:/app" `
   --memory="4g" `
   --memory-swap="8g" `
   csci316-spark-ml `
   bash -c "python3 PricePulseAI.py && jupyter nbconvert --to notebook --execute notebooks/PricePulseAI.ipynb --output-dir ./out"
'''
# %% [markdown]
# Importing some libraries

# %%
from pyspark.sql import SparkSession, functions as F, types as T

spark = SparkSession.builder.appName("CSCI316-Preprocessing").getOrCreate()

# %% [markdown]
# Reading the file

# %%
file_path = "./data/Transactions.csv"

df = (
    spark.read.option("header", "true")
    .option("inferSchema", "true")
    .csv(file_path)
)

df.show(5)

# %%
df.printSchema()
print(f"Row count: {df.count()}")

# %% [markdown]
#After running schema and row counts, I can see that the dataset has 46 columns and 1.6M rows. This is a huge dataset. I noticed that the reason there are so many columns is because some of them are duplicate columns in a different language (Arabic). So i will begin dropping the Arabic columns.
# Drop all columns ending with '_ar'
cols_to_drop = [col for col in df.columns if col.endswith("_ar")]
df = df.drop(*cols_to_drop)
print(f"Dropped {len(cols_to_drop)} columns.")

# %%
#Now I am going to remove the '_en' suffix to make it easier to deal with some of the columns
# Rename columns by removing the '_en' suffix
df = df.toDF(*[c.replace("_en", "") for c in df.columns])

# %%
#To keep things consistent in Spark, I will cast some of these columns to string.
# Cast columns to string for consistency
cols_to_categorize = ["trans_group", "property_type", "property_usage", "reg_type", "area_name"]
for col in cols_to_categorize:
    if col in df.columns:
        df = df.withColumn(col, F.col(col).cast("string"))

# %% [markdown]
#I noticed there's a lot of redundant columns that have both the Name and the ID. Since we are doing data analysis we are going to keep the names and drop the IDs to declutter the table
# Drop ID columns, keeping names
ids_to_drop = ["procedure_id", "trans_group_id", "property_type_id", "property_sub_type_id", "reg_type_id", "area_id"]
df = df.drop(*ids_to_drop)

# %%
#Now im just going to check some unique values
# Check unique values
if "rooms" in df.columns:
    df.select("rooms").distinct().show(truncate=False)

if "has_parking" in df.columns:
    df = df.withColumn("has_parking", F.col("has_parking").cast("byte"))

# %%
#I will change the type of 'has_parking' to save some space
# Simplify procedures into groups 
if "procedure_name" in df.columns:
    df = df.withColumn("procedure_name", F.col("procedure_name").cast("string"))
    proc = F.lower(F.col("procedure_name"))
    df = df.withColumn(
        "procedure_group",
        F.when(proc.contains("sell") | proc.contains("sale"), "Sales")
        .when(proc.contains("mortgage"), "Mortgage")
        .when(proc.contains("grant"), "Grants")
        .when(proc.contains("lease"), "Lease-to-Own")
        .otherwise("Other")
    )
    #I will keep procedure_name as a string column as well.
    #Then I will simplify the procedures and put them in a new column called 'procedure_group' so it can be used for analysis
    df = df.withColumn("procedure_group", F.col("procedure_group").cast("string"))
    df = df.drop("procedure_name")

if "trans_group" in df.columns:
    df = df.drop("trans_group")

# %%
# Parse dates 
if "instance_date" in df.columns:
    df = df.withColumn(
        "instance_date_parsed",
        F.coalesce(
            F.to_date(F.col("instance_date"), "d/M/yyyy"),
            F.to_date(F.col("instance_date"), "yyyy-MM-dd")
        )
    )
    df = df.filter(F.col("instance_date_parsed").isNotNull())
    df = df.withColumn("instance_date", F.col("instance_date_parsed")).drop("instance_date_parsed")

# %%
# Handle property_sub_type nulls
df = df.fillna({"property_sub_type": "Other"})
df = df.withColumn("property_sub_type", F.col("property_sub_type").cast("string"))

# %%
# Handle property_usage
if "property_usage" in df.columns:
    usage_lower = F.lower(F.trim(F.col("property_usage").cast("string")))
    df = df.withColumn("is_residential", usage_lower.contains("residential").cast("int"))
    df = df.withColumn("is_commercial", usage_lower.contains("commercial").cast("int"))
    df = df.withColumn("is_industrial", usage_lower.contains("industrial").cast("int"))
    df = df.withColumn("is_hospitality", usage_lower.contains("hospitality").cast("int"))
    df = df.withColumn("is_storage", usage_lower.contains("storage").cast("int"))
    df = df.withColumn("is_agricultural", usage_lower.contains("agricultural").cast("int"))
    df = df.withColumn("is_other_usage", (usage_lower.contains("other") | 
        ((F.col("is_residential") + F.col("is_commercial") + F.col("is_industrial")) == 0)).cast("int"))
    df = df.drop("property_usage")

# %%
# Drop redundant columns
for col_name in ["building_name", "project_number", "project_name"]:
    if col_name in df.columns:
        df = df.drop(col_name)

# %%
# Handle master_project and other fills
df = df.fillna({"master_project": "Standalone", "nearest_landmark": "No Landmark", "nearest_metro": "No Metro Access", "nearest_mall": "No Major Mall"})

# %%
# Handle rooms/bedrooms
if "rooms" in df.columns:
    if "bedroom_count" in df.columns:
        df = df.drop("bedroom_count")

    s = F.upper(F.trim(F.col("rooms").cast("string")))
    df = df.withColumn("bedrooms_raw", F.regexp_extract(s, r"(\d+)\s*B/R", 1))
    df = df.withColumn(
        "bedrooms",
        F.when(s == "STUDIO", F.lit(0))
        .when(s == "SINGLE ROOM", F.lit(1))
        .when(F.col("bedrooms_raw") != "", F.col("bedrooms_raw").cast("int"))
        .otherwise(F.lit(None).cast("int")),
    )
    df = df.withColumn("is_penthouse", (s == "PENTHOUSE").cast("byte"))
    #Your bedrooms column means number of bedrooms. For Office / Shop / GYM, bedrooms do not exist.
    df = df.drop("rooms", "bedrooms_raw")
else:
    print("Column 'rooms' not found. Logic skipped.")

# %%
#Our target column is 'actual_worth' column
# Drop redundant and party columns
redundant_cols = ["rent_value", "meter_rent_price", "no_of_parties_role_1", "no_of_parties_role_2", "no_of_parties_role_3", "transaction_id", "meter_sale_price"]
df = df.drop(*redundant_cols)

# %%
# Extract year and quarter
if "instance_date" in df.columns:
    df = df.withColumn("sale_year", F.year("instance_date"))
    df = df.withColumn("sale_quarter", F.quarter("instance_date"))
    df = df.drop("instance_date")

# %%
# Log transform target
max_val = df.select(F.max("actual_worth")).first()[0]
if max_val is not None and max_val > 50:
    df = df.withColumn("actual_worth", F.log1p(F.col("actual_worth").cast("double")))
    print("Log transform applied.")

# %%
output_path = "./data/Transactions_copy.csv"
df.coalesce(1).write.mode("overwrite").option("header", "true").csv(output_path)
print(f"Saved cleaned data to: {output_path}")