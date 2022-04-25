import re
text_all = " jg jh gjg jg jg jh   SID:004711 ilkh kh k hkj hk hkj hkj hkj hk h"
#re.match(r'\d', '3')
z = re.findall('SID:(\d\d\d\d\d\d)', text_all)
print(z[0])