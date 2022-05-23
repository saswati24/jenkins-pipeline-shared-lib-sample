#!/usr/bin/env groovy

// https://mvnrepository.com/artifact/org.apache.poi/poi
// @Grapes(
//     @Grab(group='org.apache.poi', module='poi', version='5.2.2')
// )

def call(body) {
    def config = [:]
    body.resolveStrategy = Closure.DELEGATE_FIRST
    body.delegate = config
    body()

    echo config.name
    echo "Param1 is: ${env.param1}"
    echo "Param2 is: ${env.param2}"
    if (env.param1 == 'One default') {
        echo "Param1 is default"
    }
    return this
}
