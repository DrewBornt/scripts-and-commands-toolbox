find <root directory> -type <d/f> -name "<filename>" -user <username> -group <groupname> -perm <444, u=s/u=r, etc> -size <<-> xc/xk/xM>
# -iname works like -name but is case-sensitive, x in xc/xk/xM is a variable, c is bytes, k is kilobytes and - denotes less than size
# wildcards can be used in the -name for *.txt or *.sh, etc
#
# Examples:
# 
# find / -type f -name "example.txt"
#
# to keep terminal from getting cluttered, write standard error to /dev/null with 2>/dev/null
# Example: find / -type f -name "info" 2>/dev/null
#
#
# -exec can be added to have the find command execute a command