// GroovyExcelParser
def call(body) {
    echo "Start GroovyExcelParser"

    new GroovyExcelParser(script:this).run()

    echo "Excel Read"
    currentBuild.result = 'SUCCESS' //FAILURE to fail

    return this
}
