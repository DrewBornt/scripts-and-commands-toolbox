apt-key adv --keyserver hkp://keys.gnupg.net --recv-keys <key error given mine was ED444FF07D8D0BF6>

# then add a file named <distro>.list this to your /etc/apt/sources.list.d/
# in that file add the following line

deb http://http.kali.org/kali kali-rolling main contrib no-free
deb-src http://http.kali.org/kali kali-rolling main contrib no-free