#!bin/bash

if [-d "env"]
then
    echo "Python virtual env exists"
else
    python3 -m venv env

fi

echo $PWD

source env/bin/activate

pip3 install requirements.txt

if [-d "logs"]
then
    echo "Logs folder exists"
else
    mkdir logs
    touch log/error.log log/access.log
fi    

 sudo chmod -R 777 logs   


