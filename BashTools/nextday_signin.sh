for ((i=0;i<$2;i++))

do
        echo $i
        date -s `date -d'+1 day' +%Y%m%d`
#       sleep 1s

        date_now=`date +%Y-%m-%d`

        python signin.py $date_now $1


done