@Grapes(
   [ @Grab(group='org.apache.poi', module='poi', version='3.8'),
    @Grab(group='org.apache.poi', module='poi-ooxml', version='3.8'),
   @Grab(group = 'org.apache.commons', module = 'commons-lang3', version = '3.6')]
)
import org.apache.commons.lang3.StringUtils
import org.apache.poi.ss.usermodel.*
import org.apache.poi.hssf.usermodel.*
import org.apache.poi.xssf.usermodel.*
import org.apache.poi.ss.util.*
import org.apache.poi.ss.usermodel.*
import java.io.*

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
