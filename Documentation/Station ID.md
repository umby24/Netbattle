Netbattle Station ID (SID)
1. Gets a value using 'virtual.drv', a dll that derives a unique system identifier based on your hard drive's hardware ID.
2. Makes an MD5 Hash of that value.
3. Converts each individual character in the hash into a 4 byte long hex value.
4. Cuts down whatever the length of the total thing is down to only 100 characters
5. Converts the binary to string.


SID “DeCompression”
1. Convert string into binary, keeping only the leftmost 100 bits.
 - Each CHR Value (each byte) broken into bits.
2. Take every 20th bit (19, if 0 index) into a string, and put that at the front.
3. Break up the built string in sets of 5, and turn those binary bits into a value, then add to the value as accoring to below.
4/5: if 'Y' is bigger than 8, add 56. If smaller, add 49.
