print "hello world"
foldername = "C:\\Program Files (x86)\\Jenkins\\workspace\\TestSpecifications\\33.98.0094\\data\\"

for i in range(1,3):
    filename = foldername + "testspec" + str(i) + ".txt"
    my_file = open(filename,"w+")
    #my_file.write("Hey This text is going to be added to the text file yipeee!!!")
