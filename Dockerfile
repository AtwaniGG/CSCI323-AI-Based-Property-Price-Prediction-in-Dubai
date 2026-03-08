FROM spark:3.5.0

USER root

RUN apt-get update && apt-get install -y python3 python3-pip && rm -rf /var/lib/apt/lists/*
RUN pip3 install --no-cache-dir --upgrade pip

WORKDIR /app

COPY requirements.txt /app/requirements.txt
RUN pip3 install --no-cache-dir -r /app/requirements.txt
RUN pip3 install --no-cache-dir jupyter nbconvert

COPY notebooks/ /app/notebooks/
COPY data/ /app/data/
COPY run_pipeline.sh /app/run_pipeline.sh

RUN mkdir -p /app/out

ENV PYSPARK_PYTHON=python3

CMD ["/app/run_pipeline.sh"]

