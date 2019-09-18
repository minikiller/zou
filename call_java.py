import jnius_config
import os

# #
# # # jnius_config.add_options('-Xrs', '-Xmx4096')
jnius_config.set_classpath('.', os.getcwd() + '\\*')

from jnius import autoclass

main = autoclass('com.kalix.sunlf.Main')
main1 = main()
main1.printme()
