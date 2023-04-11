# ReleasePkg.py

ReleasePkg.py is a Python script designed to assist with the creation of release packages for Intel and AMD platforms. It was created by Kevin Liou and can be used to streamline the release process.

## Features

- Supports Intel G4, G5, G6, G8, G9, G10 and AMD G4, G5, G6, G8 platforms
- Allows for input of OldVersion, NewVersion and NewBuildID
- Can process multiple projects at once
- Modifies Pkg Update Version and release note
- Auto copy Fv folder files to new release pkg

## Installation

1. Clone the repository: `git clone https://github.com/yourusername/releasepkg.git`
2. Install Python 3.x or later: https://www.python.org/downloads/
3. Install required libraries: `pip install -r requirements.txt`

## Usage

1. Move to the folder with the fv folder and the old release pkg
1. Open a terminal or command prompt
2. Navigate to the directory containing ReleasePkg.py
3. Run the script: `python ReleasePkg.py`, Or just run ReleasePkg.exe
4. Follow the prompts to input OldVersion, NewVersion and NewBuildID
5. Enter the projects you wish to process (can select multiple)
6. Wait for the script to finish processing

## Contributing

If you would like to contribute to this project, please follow these steps:

1. Fork the repository
2. Create a new branch for your feature: `git checkout -b feature-name`
3. Make changes and commit them: `git commit -am 'Add some feature'`
4. Push to the branch: `git push origin feature-name`
5. Submit a pull request
