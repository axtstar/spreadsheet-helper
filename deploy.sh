for e in $(cat .env); 
do 
    cat clasp | sed 's/\[projectid\]/'$e'/g' > .clasp.json
    npm run push
done
