FROM python:3.12-slim

RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice fontconfig && \
    apt-get clean

WORKDIR /app

COPY . .

COPY fonts /usr/share/fonts/truetype/msttcorefonts/

RUN fc-cache -f -v

RUN pip install --no-cache-dir -r requirements.txt

EXPOSE 8501

ENTRYPOINT [ "./entrypoint.sh" ]