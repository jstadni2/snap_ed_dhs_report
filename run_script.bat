:: Build the Docker image for snap_ed_dhs_report.py
docker build -t il_fcs/snap_ed_dhs_report:latest .
:: Create and start the Docker container
docker run --name snap_ed_dhs_report il_fcs/snap_ed_dhs_report:latest
:: Copy /example_outputs from the container to the build context
docker cp snap_ed_dhs_report:/snap_ed_dhs_report/example_outputs/ ./
:: Remove the container
docker rm snap_ed_dhs_report
pause