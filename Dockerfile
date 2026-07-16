FROM python:3.11-slim
WORKDIR /app

# Install dependencies
COPY excelGenerator/requirements.txt /app/
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY excelGenerator/ /app/

CMD ["bash"]
