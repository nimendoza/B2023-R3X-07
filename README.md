# B2023-R3X-07
Automation tools for the PSHS-MC Scheduling Committee

## Getting started

1. While having the folder open in admin command line, create and activate a virtual environment

```bash
python3.10 -m venv venv
. venv/bin/activate      # If on Unix
./venv/Scripts/activate  # If on Windows
```

2. Install the required packages

```bash
pip install -r requirements.txt
```

3. Install the project

```bash
pip install -e .
```

4. Create the output folder

```bash
mkdir output
```

5. After replacing the data in the input folder as needed, run the following:

```bash
python3.10 ./src/__init__.py
```
