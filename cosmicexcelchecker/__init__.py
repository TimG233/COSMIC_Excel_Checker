# This is the project for checking excels about cosmic related projects/folders

from typing import NamedTuple, Literal
class Version(NamedTuple):
    major: int
    minor: int
    micro: int
    releasetype: Literal['alpha', 'beta', 'stable']
    serial: int

version : Version = Version(major=0, minor=2, micro=0, releasetype='alpha', serial=0)
