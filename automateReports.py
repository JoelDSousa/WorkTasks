
# isto constitui um teste para verificar que devolve o número do cliente

import re

# Testing the regex used to get client number
#regex = r"(C.*)-"

#test_str = "C00238 - CMD LOUROSA"

#matches = re.search(regex,test_str)

#print(matches.group(1).strip())

def get_client_number(full_name):
    regex = r"(C.*)-"
    matches= re.search(regex,full_name)
    if matches.group(1).strip():
        return matches.group(1).strip()
    return 0



    
