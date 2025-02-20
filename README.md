# Word2Csv

Word2Csv is a Flask-based web application that allows users to upload Word documents and convert them into CSV files. The application is containerized using Docker and can be deployed on Google Cloud Run.

## Features

- Upload Word documents (.docx) and convert them to CSV format.
- Automatic cleanup of old files using APScheduler.
- Rate limiting to prevent abuse of the service.
- Secure file uploads with CSRF protection.
- Configurable environment variables for easy deployment.

## Prerequisites

- Python 3.7+
- Docker
- Google Cloud SDK (for deployment)
- A Google Cloud account with billing enabled

## Installation

1. Clone the repository:

   ```bash
   git clone https://github.com/yourusername/Word2Csv.git
   cd Word2Csv
   ```

2. Install the required Python packages:

   ```bash
   pip install -r requirements.txt
   ```

3. Set up environment variables:

   Create a `.env` file in the project root with the following variables:

   ```plaintext
   SECRET_KEY=your-secret-key
   UPLOAD_FOLDER=uploads
   ALLOWED_EXTENSIONS=docx
   FILE_RETENTION_PERIOD=3600  # in seconds
   ```

## Usage

1. Run the application locally:

   ```bash
   python app.py
   ```

2. Access the application in your web browser at `http://localhost:8080`.

3. Upload a Word document and download the converted CSV file.

## Deployment

To deploy the application on Google Cloud Run:

1. Build the Docker image:

   ```bash
   docker build -t gcr.io/[PROJECT-ID]/word2csv .
   ```

2. Push the Docker image to Google Container Registry:

   ```bash
   docker push gcr.io/[PROJECT-ID]/word2csv
   ```

3. Deploy to Google Cloud Run:

   ```bash
   gcloud run deploy word2csv \
     --image gcr.io/[PROJECT-ID]/word2csv \
     --platform managed \
     --region [REGION] \
     --allow-unauthenticated
   ```

## Contributing

Contributions are welcome! Please fork the repository and submit a pull request for any improvements or bug fixes.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Acknowledgments

- [Flask](https://flask.palletsprojects.com/) for the web framework.
- [APScheduler](https://apscheduler.readthedocs.io/) for scheduling tasks.
- [Google Cloud](https://cloud.google.com/) for cloud deployment.

```

### Key Sections Explained:

- **Features**: Highlights the main functionalities of the application.
- **Prerequisites**: Lists the tools and accounts needed to run and deploy the application.
- **Installation**: Provides step-by-step instructions to set up the project locally.
- **Usage**: Explains how to run the application and use its features.
- **Deployment**: Details the process of deploying the application to Google Cloud Run.
- **Contributing**: Encourages contributions and explains how to contribute.
- **License**: States the licensing terms for the project.
- **Acknowledgments**: Credits the tools and platforms used in the project.

Feel free to customize the README to better fit your project's specifics, such as adding more detailed instructions or additional sections as needed.
