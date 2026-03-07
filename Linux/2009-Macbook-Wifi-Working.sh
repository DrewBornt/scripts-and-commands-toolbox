# sudo all of these commands if not root
# These got wifi working on the 2009 Macbook (with Kali Linux) which did not have working wifi after Kali install. Maybe could try the bottom two commands without using the first next time if applicable.

apt-get install linux-image-$(uname -r|sed 's,[^-]*-[^-]*-,,') linux-headers-$(uname -r|sed 's,[^-]*-[^-]*-,,') broadcom-sta-dkms

modprobe -r b44 b43 b43legacy ssb brcmsmac bcma

modprobe wl