# On the destination server
Install-WindowsFeature Migration


# After Running Server Migration Tools

Send-SmigServerData -ComputerName {destination-computer-name} -SourcePath {path-to-source-folder} -DestinationPath {path-to-destination-folder} -include all -recurse

Receive-SmigServerData