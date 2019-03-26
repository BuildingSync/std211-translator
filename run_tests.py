# BuildingSync(R), Copyright (c) 2015-2018, Alliance for Sustainable Energy, LLC.
# All rights reserved.
#
# Redistribution and use in source and binary forms, with or without modification,
# are permitted provided that the following conditions are met:
#
# (1) Redistributions of source code must retain the above copyright notice, this
#     list of conditions and the following disclaimer.
#
# (2) Redistributions in binary form must reproduce the above copyright notice,
#     this list of conditions and the following disclaimer in the documentation
#     and/or other materials provided with the distribution.
#
# (3) Neither the name of the copyright holder nor the names of any contributors
#     may be used to endorse or promote products derived from this software
#     without specific prior written permission from the respective party.
#
# (4) Other than as required in clauses (1) and (2), distributions in any form of
#     modifications or other derivative works may not use the "BuildingSync"
#     trademark or any other confusingly similar designation without specific
#     prior written permission from Alliance for Sustainable Energy, LLC.
#
# THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDER(S) AND ANY CONTRIBUTORS "AS
# IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
# IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
# DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER(S), ANY CONTRIBUTORS, THE
# UNITED STATES GOVERNMENT, OR THE UNITED STATES DEPARTMENT OF ENERGY, NOR ANY
# OF THEIR EMPLOYEES, BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL,
# EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO,
# PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR
# BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER
# IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
# ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
# POSSIBILITY OF SUCH DAMAGE.

import read211
import loadxl
import warnings
import urllib.request
from lxml import etree
from io import StringIO

#remote_file = urllib.request.urlopen('https://github.com/BuildingSync/schema/blob/develop/BuildingSync.xsd')
#schema_data = remote_file.read()

#doc = etree.parse(StringIO(schema_data.decode()))
fp = open('../bsxml/BuildingSync.xsd', 'r')
txt = fp.read()
fp.close()
#tree = etree.parse('https://github.com/BuildingSync/schema/blob/develop/BuildingSync.xsd')
tree = etree.parse('../bsxml/BuildingSync.xsd')
#doc = etree.parse(txt)
schema = etree.XMLSchema(tree)

legit = etree.parse('../bsxml/examples/Golden Test File.xml')

test_files = ['examples/std211_example.xlsx']

for file in test_files:
    warnings.simplefilter("ignore")
    wb = loadxl.load_workbook(file)
    std211 = read211.read_std211_xls(wb)
    bsxml = read211.map_to_buildingsync(std211, groupspaces=True)
    try:
        schema.assertValid(bsxml)
    except etree.DocumentInvalid as exc:
        print(exc)
    warnings.simplefilter("default")

