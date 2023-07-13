# This is the project for checking excels about cosmic related projects/folders

from typing import NamedTuple, Literal
class Version(NamedTuple):
    major: int
    minor: int
    micro: int
    releasetype: Literal['alpha', 'beta', 'stable']
    serial: int

version : Version = Version(major=0, minor=1, micro=3, releasetype='stable', serial=0)
