#!/bin/bash

function pause(){
   read -p "$*"
}
javac -cp ".;lib/json-org-20140107.jar;lib/jacob-1.18.jar;lib/jsoup-1.15.3.jar" Main.java

pause 'Press [Enter] key to continue...'