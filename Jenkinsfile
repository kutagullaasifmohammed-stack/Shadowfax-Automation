pipeline {
    agent any

    stages {
        stage('Checkout') {
            steps {
                echo 'Checking out project files...'
                checkout scm
            }
        }

        stage('Compile Project') {
            steps {
                echo 'Compiling Java code via Maven...'
                bat 'mvn clean compile'
            }
        }

        stage('Execute Selenium Script') {
            steps {
                echo 'Starting E-Way Bill Automation...'
                echo 'ATTENTION: Keep Chrome open to enter CAPTCHA/OTP manually when prompted!'
                bat 'mvn exec:java'
            }
        }
    }

    post {
        always {
            echo 'Pipeline execution complete.'
        }
        success {
            echo 'Excel file processed successfully.'
        }
        failure {
            echo 'Script failed! Check logs or verify if AutoExtendFile.xlsx is opened elsewhere.'
        }
    }
}