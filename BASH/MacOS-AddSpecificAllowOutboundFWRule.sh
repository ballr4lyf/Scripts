#!/bin/sh

destNetwork="1.2.3.4/24"
destPort="Any"
pfFile="myPfFileName"

# Check if the Packet Filter (pf) file exists.
if [! -f /etc/pf.anchors/pfFile]
then
    touch /etc/pf.anchors/pfFile
fi

