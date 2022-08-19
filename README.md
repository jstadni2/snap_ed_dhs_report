# SNAP-Ed DHS Report

The SNAP-Ed DHS Report script generates metrics for Illinois Extension's quarterly SNAP-Ed evaluation report. Data from [PEARS](https://www.k-state.edu/oeie/pears/) is used to calculate evaluation metrics
specified by the [SNAP-Ed Evaluation Framework](https://snapedtoolkit.org/framework/index/) and [Illinois Department of Human Services \(IDHS\)](https://www.dhs.state.il.us/page.aspx).

## Installation

The recommended way to install the SNAP-Ed DHS Report script is through git, which can be downloaded [here](https://git-scm.com/downloads). Once downloaded, run the following command:

```bash
git clone https://github.com/jstadni2/snap_ed_dhs_report
```

Alternatively, this repository can be downloaded as a zip file via this link:
[https://github.com/jstadni2/snap_ed_dhs_report/zipball/master/](https://github.com/jstadni2/snap_ed_dhs_report/zipball/master/)

This repository is designed to run out of the box on a Windows PC using Docker and the [/example_inputs](https://github.com/jstadni2/snap_ed_dhs_report/tree/master/example_inputs) and [/example_outputs](https://github.com/jstadni2/snap_ed_dhs_report/tree/master/example_outputs) directories.
To run the script in its current configuration, follow [this link](https://docs.docker.com/desktop/windows/install/) to install Docker Desktop for Windows. 

With Docker Desktop installed, this script can be run simply by double-clicking the `run_script.bat` file in your local directory.

The `run_script.bat` file can also be run in Command Prompt by entering the following command with the appropriate path:

```bash
C:\path\to\snap_ed_dhs_report\run_script.bat
```

### Setup instructions for SNAP-Ed implementing agencies

The following steps are required to execute the SNAP-Ed DHS Report script using your organization's PEARS data:
1. Contact [PEARS support](mailto:support@pears.io) to set up an [AWS S3](https://aws.amazon.com/s3/) bucket to store automated PEARS exports.
2. Download the automated PEARS exports. Illinois Extension's method for downloading exports from the S3 is detailed in the [PEARS Nightly Export Reformatting script](https://github.com/jstadni2/pears_nightly_export_reformatting/blob/6f370389776fb8f88495fbe4e7918c203fd84997/pears_nightly_export_reformatting.py#L9-L45).
3. Set the appropriate input and output paths in `snap_ed_dhs_report.py` and `run_script.bat`.
	- The [Input Files](#input-files) and [Output Files](#output-files) sections provide an overview of required and output data files.
	- Copying input files to the build context would enable continued use of Docker and `run_script.bat` with minimal modifications.

### Additional setup considerations

- The formatting of PEARS export workbooks changes periodically. The example PEARS exports included in the [/example_inputs](https://github.com/jstadni2/snap_ed_dhs_report/tree/master/example_inputs) directory are based on workbooks downloaded on 08/12/22.
Modifications to `snap_ed_dhs_report.py` may be necessary to run with subsequent PEARS exports.
- Illinois Extension utilized [Task Scheduler](https://docs.microsoft.com/en-us/windows/win32/taskschd/task-scheduler-start-page) to run this script from a Windows PC on a quarterly basis.
- Plans to deploy the SNAP-Ed DHS Report script on AWS were never implemented and are currently beyond the scope of this repository.
- Other SNAP-Ed implementing agencies intending to utilize the SNAP-Ed DHS Report script should consider the following adjustments as they pertain to their organization:
	- The metrics provided in this report are based in part on guidance from IDHS. Other regulatory agencies overseeing SNAP-Ed may require additional evaluation metrics.
	
## Input Files

The following input files are required to run the SNAP-Ed DHS Report script:
- Reformatted PEARS module exports output from the [PEARS Nightly Export Reformatting script](https://github.com/jstadni2/pears_nightly_export_reformatting):
    - [Coalition_Export.xlsx](https://github.com/jstadni2/snap_ed_dhs_report/blob/master/example_inputs/Coalition_Export.xlsx)
    - [Indirect_Activity_Export.xlsx](https://github.com/jstadni2/snap_ed_dhs_report/blob/master/example_inputs/Indirect_Activity_Export.xlsx)
    - [Partnership_Export.xlsx](https://github.com/jstadni2/snap_ed_dhs_report/blob/master/example_inputs/Partnership_Export.xlsx)
    - [Program_Activities_Export.xlsx](https://github.com/jstadni2/snap_ed_dhs_report/blob/master/example_inputs/Program_Activities_Export.xlsx)
    - [PSE_Site_Activity_Export.xlsx](https://github.com/jstadni2/snap_ed_dhs_report/blob/master/example_inputs/PSE_Site_Activity_Export.xlsx)

Example input files are provided in the [/example_inputs](https://github.com/jstadni2/snap_ed_dhs_report/tree/master/example_inputs) directory. 

## Output Files

The following output files are produced by the SNAP-Ed DHS Report script:
- [DHS Report FY2022 FQ.xlsx](https://github.com/jstadni2/snap_ed_dhs_report/blob/master/example_outputs/DHS%20Report%20FY2022%20Q3.xlsx): A workbook that compiles the quarterly SNAP-Ed evaluation metrics reported to IDHS.

Example output files are provided in the [/example_outputs](https://github.com/jstadni2/snap_ed_dhs_report/tree/master/example_outputs) directory.