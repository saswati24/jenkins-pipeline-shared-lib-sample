@Grapes(
   [ @Grab(group='org.apache.poi', module='poi', version='3.8'),
    @Grab(group='org.apache.poi', module='poi-ooxml', version='3.8'),
   @Grab(group = 'org.apache.commons', module = 'commons-lang3', version = '3.6')]
)
import org.apache.commons.lang3.StringUtils

class Deployer {
    int tries = 0
    Script script

    def run() {
        while (tries < 10) {
            Thread.sleep(1000)
            tries++
            script.echo("tries is numeric: " + StringUtils.isAlphanumeric("" + tries))
        }
    }
}
