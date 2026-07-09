pipeline {
    agent {
    kubernetes {
        yaml '''
apiVersion: v1
kind: Pod
spec:
  containers:
    - name: kaniko
      image: gcr.io/kaniko-project/executor:v1.23.2-debug
      command:
        - /busybox/sh
      args:
        - -c
        - sleep infinity
      tty: true
      volumeMounts:
        - name: docker-config
          mountPath: /kaniko/.docker

    - name: gitops
      image: alpine:3.21
      command:
        - /bin/sh
      args:
        - -c
        - sleep infinity
      tty: true

  volumes:
    - name: docker-config
      secret:
        secretName: dockerhub-secret
        items:
          - key: .dockerconfigjson
            path: config.json
'''
    }
}

    environment {
        IMAGE_NAME = 'docker.io/inteligeninfosys/demand-letter-api-stima'

        DEPLOY_REPO_URL = 'https://github.com/inteligeninfosys/ecollect-deployments.git'
        DEPLOY_REPO_BRANCH = 'stima_deploy'
        DEPLOY_REPO_DIRECTORY = 'deployment-repo'

        DEPLOY_APP = 'demand-letter-api'
        DEPLOY_ENVIRONMENT = 'prod'
        DEPLOY_KUSTOMIZATION = 'demand-letter-api/overlays/prod/kustomization.yaml'
    }

    stages {
        stage('Checkout') {
            steps {
                git(
                    branch: 'main',
                    credentialsId: 'github_credentials',
                    url: 'https://github.com/inteligeninfosys/demand-letter-api.git'
                )
            }
        }

        stage('Generate Image Tags') {
            steps {
                script {
                    env.IMAGE_TAG = sh(
                        script: 'date +%Y%m%d-%H%M%S',
                        returnStdout: true
                    ).trim() + "-${env.BUILD_NUMBER}"

                    env.GIT_SHA = sh(
                        script: 'git rev-parse --short HEAD',
                        returnStdout: true
                    ).trim()

                    echo "IMAGE_TAG=${env.IMAGE_TAG}"
                    echo "GIT_SHA=${env.GIT_SHA}"
                }
            }
        }

        stage('Build and Push') {
            steps {
                container('kaniko') {
                    sh '''
                        set -eu

                        /kaniko/executor \
                          --context="${WORKSPACE}" \
                          --dockerfile="${WORKSPACE}/Dockerfile" \
                          --destination="${IMAGE_NAME}:${IMAGE_TAG}" \
                          --destination="${IMAGE_NAME}:${GIT_SHA}" \
                          --destination="${IMAGE_NAME}:latest"

                        echo "Kaniko build and push completed"
                    '''
                }
            }
        }

        stage('Update Deployment Repository') {
            steps {
                container('gitops') {
                    withCredentials([
                        gitUsernamePassword(
                            credentialsId: 'github_credentials',
                            gitToolName: 'Default'
                        )
                    ]) {
                        sh '''
                            set -eu

                            apk add --no-cache git curl yq

                            rm -rf "${DEPLOY_REPO_DIRECTORY}"

                            git clone \
                            --branch "${DEPLOY_REPO_BRANCH}" \
                            "${DEPLOY_REPO_URL}" \
                            "${DEPLOY_REPO_DIRECTORY}"

                            cd "${DEPLOY_REPO_DIRECTORY}"

                            UPDATE_BRANCH="deploy/${DEPLOY_APP}-${DEPLOY_ENVIRONMENT}-${IMAGE_TAG}"

                            git checkout -b "${UPDATE_BRANCH}"

                            export NEW_IMAGE_TAG="${IMAGE_TAG}"

                            yq -i '
                            (.images[] |
                            select(.name == "docker.io/inteligeninfosys/demand-letter-api-stima") |
                            .newTag) = strenv(NEW_IMAGE_TAG)
                            ' "${DEPLOY_KUSTOMIZATION}"

                            echo "Updated deployment manifest:"
                            cat "${DEPLOY_KUSTOMIZATION}"

                            git config user.name "Jenkins CI"
                            git config user.email "jenkins@inteligen.co.ke"

                            git add "${DEPLOY_KUSTOMIZATION}"

                            if git diff --cached --quiet; then
                                echo "No deployment manifest changes detected"
                                exit 0
                            fi

                            git commit \
                            -m "deploy(demand-letter-api): update prod image to ${IMAGE_TAG}"

                            git push origin "${UPDATE_BRANCH}"

                            echo "${UPDATE_BRANCH}" > ../deployment-branch.txt
                        '''
                    }
                }
            }
        }

        stage('Create Deployment Pull Request') {
            steps {
                container('gitops') {
                    withCredentials([
                        usernamePassword(
                            credentialsId: 'github_credentials',
                            usernameVariable: 'GITHUB_USERNAME',
                            passwordVariable: 'GITHUB_TOKEN'
                        )
                    ]) {
                        sh '''
                            set -eu

                            apk add --no-cache curl jq

                            UPDATE_BRANCH="$(cat deployment-branch.txt)"

                            PR_TITLE="Deploy demand-letter-api PROD ${IMAGE_TAG}"

                            PR_BODY="Automated deployment update from Jenkins.

        Application: demand-letter-api
        Environment: PROD
        Image: ${IMAGE_NAME}:${IMAGE_TAG}
        Git SHA: ${GIT_SHA}
        Jenkins build: ${BUILD_NUMBER}"

                            API_RESPONSE=$(curl --fail-with-body \
                            --silent \
                            --show-error \
                            --request POST \
                            --header "Authorization: Bearer ${GITHUB_TOKEN}" \
                            --header "Accept: application/vnd.github+json" \
                            --header "X-GitHub-Api-Version: 2022-11-28" \
                            "https://api.github.com/repos/inteligeninfosys/ecollect-deployments/pulls" \
                            --data "$(jq -n \
                                --arg title "${PR_TITLE}" \
                                --arg head "${UPDATE_BRANCH}" \
                                --arg base "${DEPLOY_REPO_BRANCH}" \
                                --arg body "${PR_BODY}" \
                                '{
                                title: $title,
                                head: $head,
                                base: $base,
                                body: $body
                                }'
                            )"
                            )

                            PR_URL=$(echo "${API_RESPONSE}" | jq -r '.html_url')

                            if [ -z "${PR_URL}" ] || [ "${PR_URL}" = "null" ]; then
                                echo "GitHub did not return a pull request URL"
                                echo "${API_RESPONSE}"
                                exit 1
                            fi

                            echo "Deployment pull request created: ${PR_URL}"
                            echo "${PR_URL}" > deployment-pr-url.txt
                        '''
                    }
                }
            }
        }
    }

    post {
        success {
            echo "Successfully pushed:"
            echo "${IMAGE_NAME}:${IMAGE_TAG}"
            echo "${IMAGE_NAME}:${GIT_SHA}"
            echo "${IMAGE_NAME}:latest"
        }

        failure {
            echo 'Image build or push failed.'
        }
    }
}
