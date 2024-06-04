async function get_sheet(){

    var getSheetURL = 'https://chase35x.pythonanywhere.com/make_sheet?param=AAPL'

    let response = await fetch(getSheetURL)
        .then(data => {
            return data;
        })           //api for the get request
    
    
    console.log(response)

}