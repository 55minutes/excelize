from ConfigParser import SafeConfigParser
import os


def get_production_version():
    version_config = SafeConfigParser()
    version_config.readfp(open(
        os.path.join(os.path.dirname(__file__), 'VERSION.cfg')))
    return version_config.get('version', 'production')
