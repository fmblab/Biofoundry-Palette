# Biofoundry palette
<p align = "center">
  <img src="https://lh6.googleusercontent.com/VekUq4tnfzO4vXNuJ7vPwTQcfIQSrxwENfGXRKKwdtDGWLTLwFsMjaqm1XEbHBTxZ7bso1ksD0bbp_BKH36bmL4=w16383"/>
 </p>

Biofoundry Palette is a python-based planning-assistance software for synthetic biology applications in an automated workflow and biofoundry facilities, enabling automated high-throughput experiments. The program in this repository is for liquid handler-based experimentation and operation in the biofoundry workflow. Step 6~11 is still undergoing in volume and positional optimization. Further update will be uploaded as soon as possible. If you want to get new on the update, please contact the [BioFoundry Research Center](https://swb.skku.edu/BioFoundryRC/index.do).

## Installation and Execution
Biofoundry Palette Program does not require any additional installation. In order to execute the file, a please click BFP.exe

**Operation Instruction**
- OPERATOR: place the name of the person running the machine
- USER: place the name of the customer
- ORDER ID: place the name of ther order identification
- ORDER VOLUME: place the amount of the order
- NEW/OLD WELL: choice of using old/used well
- PLATE ID: place the identification of old/used well
- WELL POSITION: starting column position of the well
- STEP: step 1-5
- ORDER DATE: date of order is placed
- RUN DATE: date of product to be sent
- FILE NAME: name of the instruction file

However, when if one wants to run using raw python script, please install the following python packages

**Pre-requisites:**
The following commands is to be executed in Windows Powershell
<br/>
[openpyxl](https://openpyxl.readthedocs.io/en/stable/)
| OS | Command |
| ----------- | ----------- |
| Windows | pip install openpyxl |

[xlsxwriter](https://xlsxwriter.readthedocs.io/)
| OS | Command |
| ----------- | ----------- |
| Windows | pip install xlsxwriter |

[tkcalendar](https://pypi.org/project/tkcalendar/)
| OS | Command |
| ----------- | ----------- |
| Windows | pip install tkcalendar |

[tkcalendar](https://docs.python.org/ko/3/library/tkinter.html)
| OS | Command |
| ----------- | ----------- |
| Windows | pip install tkinter |

## The following is flowchart of Biofoundry Palette
```mermaid
graph LR
A((BFP.exe))--Input Information-->B[bridge.py]
E[format.xlsx]--copy format data-->D[format_reader.py]
D[format_reader.py]--create file-->C((result.xlsx))
B[bridge.py]--instruct to create destination file-->D[format_reader.py]
B[bridge.py]--send all information-->F[adder.py]
F[adder.py]--Insert order information-->C((result.xlsx))
F[adder.py]--Order Volume-->G[position.py]
G[position.py]--insert information-->C((result.xlsx))
F[adder.py]--Order Volume-->H[volume.py]
H[volume.py]--insert information-->C((result.xlsx))
F[adder.py]--Order Information-->I[save_data.py]
D[format_reader.py]--create file-->J((past_order.xlsx))
I[save_data.py]--insert information-->J((past_order.xlsx))
```
**Formulation used in position.py and volume.py**
<p align = "center">
  <img src="https://user-images.githubusercontent.com/57700611/166395788-2004a085-4ca9-4b36-a466-c148f91bd37d.png"/>
 </p>

## Publishing
The repository is set up to prohibit commits directly to the main branch. GitHub PRs must be approved by at least one other developer before they can be merged into main branch.

## Additional Information
To learn more about Biofoundry Palette, including program edittaion instructions and documentation, visit [BioFoundry Research Center](https://swb.skku.edu/BioFoundryRC/index.do).

![Step_figure](Biofoundry-Palette/Step_figure.jpg)
