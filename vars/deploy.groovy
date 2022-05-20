#!/usr/bin/env groovy

def call(String path) {
    echo "Start Deploy"

    new Deployer(script:this).run(path)

    echo "Deployed"
    currentBuild.result = 'SUCCESS' //FAILURE to fail

    return this
}
