# Excel to YAML converter

[![Build Status](https://travis-ci.org/titom73/python-excel-serializer.svg?branch=master)](https://travis-ci.org/titom73/python-excel-serializer)

Script provides a mechanism to extract information from any Excel sheet and create a text file with any structure. Transformation is based by using 

## Description

Repository introduced a way to convert data from an Excel file to YAML structure. It is convinient to share information with customers / team to collect all data. Once this process is over, python script will extract information from this excel file and create YAML outputs that can be used by ansible.

# Python Excel Reader

## Installation

To start working with this script, you have to installa dependencies with following command:

```shell
$ pip install -r requirements.txt
$ python bin/inetsix-excel-to-template -h
```

Or you can use `pip` to install script in your `$PATH`

```shell
$ pip install git+https://github.com/titom73/python-excel-serializer.git
```

## Description

__Script name:__ [`bin/inetsix-excel-to-template`](bin/inetsix-excel-to-template)

__Supported features__:

- Local file fetching.
- Remote file fetching using HTTP or HTTPS.
- Table with N columns.
- List of items.
- Jinja2 engine to render file in any text format.

```shell
python bin/inetsix-excel-to-template -h
usage: tools.python.read.excel.py [-h] [-e EXCEL] [-u URL] [-s SHEET]
                                  [-m MODE] [-n NB_COLUMNS] [-t TEMPLATE]
                                  [-o OUTPUT] [-v]

Excel to text file

optional arguments:
  -h, --help            show this help message and exit
  -e EXCEL, --excel EXCEL
                        Input Excel file
  -u URL, --url URL     URL for remote Excel file
  -s SHEET, --sheet SHEET
                        Excel sheet tab
  -m MODE, --mode MODE  Serializer mode: table / list
  -n NB_COLUMNS, --nb_columns NB_COLUMNS
                        Number of columns part of the table
  -t TEMPLATE, --template TEMPLATE
                        Template file to render
  -o OUTPUT, --output OUTPUT
                        Output file
  -v, --verbose         Increase verbositoy for debug purpose
```

This script reads a sheet (`-s`) in an excel file (`-e`) and extract data into the following python structure. If file is stored on a HTTP/HTTPS server, `-u` option will download it directly and will replace `-e` option.

```yaml
SHEET_NAME:
  - COLUMN#1: value ROW#1 / COLUMN#1
    COLUMN#2: value ROW#1 / COLUMN#2
    ...
  - COLUMN#1: value ROW#2 / COLUMN#1
    COLUMN#2: value ROW#2 / COLUMN#2
    ...
```

Since this structure is `yaml` compliant, script opens a template (`-t`) based on `jinja2` format to render output file (`-o`)

### Excel data structure description

__Table representation__:

A table is a two dimension represented like this:

| Local Device | Local Port | Local Port Name | Remote Port Name | Remote Port | Remote device |
| --- | --- | --- | --- | --- | --- | 
| poc-qfx5110-169 | port1 | et-0/0/0 | et-0/0/0 | port1 | poc-qfx5110-172 |
| poc-qfx5110-169 | port1 | et-0/0/1 | et-0/0/0 | port1 | poc-qfx5110-174 |

In this case, every column name is a key name. In term of Python representation, structure built is like this where `topology` is the name of the Excel sheet:

```python
topology:
- {id: '1', local_device: demo-qfx10k2-14, local_port: port1, local_port_name: et-0/0/0,
  remote_device: demo-qfx5110-11, remote_port: port1, remote_port_name: et-0/0/0}
- {id: '2', local_device: demo-qfx10k2-14, local_port: port2, local_port_name: et-0/0/1,
  remote_device: demo-qfx5110-12, remote_port: port1, remote_port_name: et-0/0/0}
```

__List representation__:

A list is a structure where every row is a key:

| Key | Data |
| --- | --- |
| backup_destination | 1.1.1.0/24 1.1.2.0/24 |
| backup_router | 2.2.2.2 |
| domain_name | lab.inetsix.net |
| dual_re | true |

In this case, output is like this:

```python
global:
  backup_destination: 1.1.1.0/24 1.1.2.0/24
  backup_router: 2.2.2.2
  domain_name: lab.inetsix.net
  dual_re: 'true'
```

## Usage

Below is an output example with `-v` enable:

```shell
python python-tools/tools.python.read.excel.py \
			-e examples/topology.xlsx\
			-s Topology \
			-t examples/topology.j2 \
			-o examples/output.txt -v

			
 * Reading excel file: examples/topology.xlsx
 ** Debug Output **
topology:
- {id: '1', local_device: dev1, local_port: port1, local_port_name: et-0/0/0, remote_device: dev2,
  remote_port: port1, remote_port_name: et-0/0/0}
- {id: '2', local_device: dev1, local_port: port2, local_port_name: et-0/0/1, remote_device: dev3,
  remote_port: port1, remote_port_name: et-0/0/0}
- {id: '3', local_device: dev1, local_port: port3, local_port_name: et-0/0/2, remote_device: dev4,
  remote_port: port1, remote_port_name: et-0/0/0}
- {id: '4', local_device: dev1, local_port: port4, local_port_name: et-0/0/3, remote_device: dev5,
  remote_port: port1, remote_port_name: et-0/0/0}
- {id: '5', local_device: dev1, local_port: port5, local_port_name: et-0/0/4, remote_device: dev6,
  remote_port: port1, remote_port_name: et-0/0/0}
- {id: '6', local_device: dev1, local_port: port6, local_port_name: et-0/0/5, remote_device: dev7,
  remote_port: port1, remote_port_name: et-0/0/0}
- {id: '7', local_device: dev1, local_port: port7, local_port_name: et-0/0/6, remote_device: dev8,
  remote_port: port1, remote_port_name: et-0/0/0}

 * Template rendering
    > Use template: examples/topology.j2
    > Output file: examples/output.txt
```

## Example use case

### Excel Inputs

Assuming table is the following:

| Local Device | Local Port | Local Port Name | Remote Port Name | Remote Port | Remote device |
| --- | --- | --- | --- | --- | --- | 
| poc-qfx5110-169 | port1 | et-0/0/0 | et-0/0/0 | port1 | poc-qfx5110-172 |
| poc-qfx5110-169 | port1 | et-0/0/1 | et-0/0/0 | port1 | poc-qfx5110-174 |


### Template for rendering

```jinja2
topo:
{%- for tester in topology %}
{%- if loop.previtem is not defined or tester.local_device != loop.previtem.local_device %}
  {{tester.local_device}}:
{%- for link in topology %}
{%- if tester.local_device == link.local_device %}
    {{link.local_port}}: { name: {{link.local_port_name}},    peer: {{link.remote_device}},      pport: {{link.remote_port}},     type: ebgp, link: {{link.id}},   linkend: 1 }
{%- endif %}
{%- endfor %}
{%- endif %}

{%- endfor %}
{%- for tester in topology|sort(attribute='remote_device') %}
{%- if loop.previtem is not defined or tester.remote_device != loop.previtem.remote_device %}
  {{tester.remote_device}}:
{%- for link in topology|sort(attribute='remote_device') %}
{%- if tester.remote_device == link.remote_device %}
    {{link.remote_port}}: { name: {{link.remote_port_name}},    peer: {{link.local_device}},      pport: {{link.local_port}},     type: ebgp, link: {{link.id}},   linkend: 2 }
{%- endif %}
{%- endfor %}
{%- endif %}

{%- endfor %}
```

### Output rendered

Output rendering is similar to:

```yaml
topo:
  poc-qfx5110-169:
    port1: { name: et-0/0/0,    peer: poc-qfx5110-172,      pport: port1,     type: ebgp, link: 1,   linkend: 1 }
    port2: { name: et-0/0/1,    peer: poc-qfx5110-174,      pport: port1,     type: ebgp, link: 2,   linkend: 1 }
    port3: { name: et-0/0/2,    peer: poc-qfx5110-175,      pport: port1,     type: ebgp, link: 3,   linkend: 1 }
    port4: { name: et-0/0/3,    peer: poc-qfx5110-188,      pport: port1,     type: ebgp, link: 4,   linkend: 1 }
    port5: { name: et-0/0/4,    peer: poc-qfx5110-189,      pport: port1,     type: ebgp, link: 5,   linkend: 1 }

...

  poc-qfx5110-172:
    port1: { name: et-0/0/0,    peer: poc-qfx5110-169,      pport: port1,     type: ebgp, link: 1,   linkend: 2 }
    port2: { name: et-0/0/1,    peer: poc-qfx5110-170,      pport: port1,     type: ebgp, link: 6,   linkend: 2 }
    port3: { name: et-0/0/2,    peer: poc-qfx5110-171,      pport: port1,     type: ebgp, link: 11,   linkend: 2 }

...
```