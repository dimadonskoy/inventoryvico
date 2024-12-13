@echo off

docker login -u crooper22@gmail.com

docker tag returnapp crooper/web-apps-verifone:returnapp

docker push crooper/web-apps-verifone:returnapp