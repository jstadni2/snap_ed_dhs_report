FROM python:3.9

WORKDIR /snap_ed_dhs_report

COPY . .

RUN pip install -r requirements.txt

CMD [ "python", "./snap_ed_dhs_report.py" ]