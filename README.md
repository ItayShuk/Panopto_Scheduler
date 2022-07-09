<!-- ABOUT THE PROJECT -->
## About The Project
Improvement for ILAN schedule - https://github.com/ilanerukh/Panopto-Scheduler

Schedule recording using PANOPTO API and Google form which connected to Google Project, while filling: course id, year, semester, hall, name of the lecture and whether it is repetetive lecture to the google form, the program will automatically schedule your request.
At the end of the process, the schedule will send you an email with recording details,
If the scheduling failed for some reason, it will respond with error mail.


<!-- GETTING STARTED -->
## Getting Started


### Prerequisites

Use Python 3.8, and run
```sh
pip install -r requirements.txt

```

### Installation

1. Sign in to the Panopto web site as Administarator
2. Click the System icon at the left-bottom corner.
3. Click API Clients
4. Click New
5. Enter arbitrary Client Name
6. Select Server-side Web Application type.
7. Enter https://localhost into CORS Origin URL.
8. Enter http://localhost:9127/redirect into Redirect URL.
9. The rest can be blank. Click "Create API Client" button.
10. Note the created Client ID and Client Secret.



<!-- USAGE EXAMPLES -->
## Usage
The client id and client secret are necessary. If you provide only them, all the database will be migrated.
You can add course id, year, semester. In this case only what you entered will be migrated.
In addition, you can add folder id. In this case, what you ask to upload will be uploaded to this specific panopto folder id.

In order to run, run with shell, or with Pycharm with those arguments:
```
scheduler.py --client-id <panopto client id> --client-secret <panopto client secret> --user <email username> --password <email password> --google-json <path to client secret in sheets>```

Email and password currently support outlook only, and this argument is optional.
Logins will not be saved or used for any purpose.


<!-- LICENSE -->
## License

Distributed under the MIT License. See `LICENSE` for more information.


