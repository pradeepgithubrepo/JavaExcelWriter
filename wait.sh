f() { sleep 5; return 13; }
f &
wait $!
echo $?
