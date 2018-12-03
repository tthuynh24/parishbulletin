import re
from configparser import ConfigParser
from ast import literal_eval as Eval

parm = re.compile("(\w+.\w+)|(\w+)")

def readCfg(cfgFn):
    myCfg = ConfigParser()
    with open(cfgFn, "r") as f:
        myCfg.readfp(f)
    return myCfg

def extendCfg(myCfg, cfgFiles):
    myCfg.read(cfgFiles)

class MyConfig(ConfigParser):
    def __init__(self, conffile):
        super().__init__()
        self.conf_file = conffile
        try: self.readfp(open(self.conf_file), 'r')
        except IOError as Err:
            if Err.errno == 2: pass
            else: raise Err

        # self.__delbug()

    def set(self, section, option, value):
        if self.has_section(section):
            ConfigParser.set(self, section, option, value)
        else:
            self.add_section(section)
            ConfigParser.set(self, section, option, value)

    def getOptVal(self, section, option):
        try: return Eval(parm.sub(r'"\1"', self.get(section, option)))
        except Exception: return None

    def save(self):
        self.write(open(self.conf_file, 'w'))

    def __del__(self):
        pass

    def delbug(self):
        for sectn in self.sections():
            for opt in self.options(sectn):
                val = self.getOptVal(sectn, opt)
                print('====>> MyConfig: ', sectn, opt, val)

##############################################################################
# cfg.sections()        return list of sections
# cfg.options(section)  return list of options
# End of module                                                              #
##############################################################################
