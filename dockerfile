FROM python:3.7-slim
COPY . /Sennheiser-Reporting-Tools
WORKDIR /Sennheiser-Reporting-Tools
RUN pip install -r requirements.txt
EXPOSE 80
RUN mkdir ~/.streamlit
RUN cp config.toml ~/.streamlit/config.toml
RUN cp credentials.toml ~/.streamlit/credentials.toml
WORKDIR /Sennheiser-Reporting-Tools
ENTRYPOINT ["streamlit", "run"]
CMD ["app.py"]