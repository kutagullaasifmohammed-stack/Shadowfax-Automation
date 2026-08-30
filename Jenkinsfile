pipeline {
    agent any

    tools {
        // Ensure Maven is configured in Jenkins under 'Global Tool Configuration' with the name 'Maven 3'
        maven 'Maven 3'
    }

    stages {
        stage('Checkout') {
            steps {
                git branch: 'main', url: 'https://github.com/kutagullaasifmohammed-stack/Shadowfax-Automation.git'
            }
        }

        stage('Build & Test') {
            steps {
                // Runs your TestNG suite defined in pom.xml
                sh 'mvn clean test' 
            }
        }
    }

    post {
        always {
            // Publishes TestNG results if plugin is installed
            testng testResultsPattern: '**/target/surefire-reports/testng-results.xml'
        }
    }
}
