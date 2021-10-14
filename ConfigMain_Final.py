import logic


hello = logic.display_instance(input("What is the name of the file you want to check:: ").replace(".xlsm","") + ".xlsm")
hello.configGen()


