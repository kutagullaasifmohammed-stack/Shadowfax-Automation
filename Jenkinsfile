pipeline {
    agent any

    tools {
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
                bat 'mvn clean test' 
            }
        }
    }

    post {
        always {
            junit '**/target/surefire-reports/*.xml'
        }
    }
}
