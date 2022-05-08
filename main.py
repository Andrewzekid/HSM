from MS import MultiSelect,access_website,login
if __name__=="__main__":
    #set encoding method for japanese email
    #'leonaellen@gmail.com',
    # time.sleep(3)
    # zoomOut("optionsdarkblue.jpg",number_of_zooms=4)

    #initialize multiselect object
    target_list = [(" ピコタン ","images/keyword2.jpg")]
    # target_list = [(" ポーチ 《フールビ》 20 ","images/testkeyword.jpg")]
    Multisearch = MultiSelect(target=target_list)

    access_website()
    login()

    #main loop
    Multisearch.search()

