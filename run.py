from time import time
time_start = time()

import import_oracle

data_in = import_oracle.extract('/Users/wep/code/CTSI.xlsx', 5)

result = import_oracle.store(data_in, '/Users/wep/code/column_format.xlsx', '/Users/wep/code/new_file.xlsx')

duration = time() - time_start

print round(duration,1), ' seconds to complete'