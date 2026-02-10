# 使用官方 Python 基础镜像
FROM python:3.9-slim

# 设置工作目录
WORKDIR /app

# 复制当前目录下的所有文件到容器中
COPY . .

# 安装依赖
RUN pip install -r requirements.txt

# 暴露 8080 端口
EXPOSE 8080

# 启动 Streamlit，强制使用 8080 端口
CMD ["streamlit", "run", "dashboard.py", "--server.port", "8080", "--server.address", "0.0.0.0"]
