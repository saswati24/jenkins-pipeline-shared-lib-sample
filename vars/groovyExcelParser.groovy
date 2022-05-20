// GroovyExcelParser
def call(body) {
    echo "Start GroovyExcelParser"
    
    git url: "https://github.com/saswati24/simple-java-maven-app.git"
    sh 'mvn clean install'

    new GroovyExcelParser(script:this).run()

    echo "Excel Read"
    currentBuild.result = 'SUCCESS' //FAILURE to fail

    return this
}
