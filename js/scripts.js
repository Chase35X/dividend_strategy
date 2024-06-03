
async function test(){
    const data = { key: 'value' };
    const url = 'https://cryptic-headland-94862.herokuapp.com/http://127.0.0.1:5000/api/test'; // Replace with your public IP address

    console.log('Sending data to:', url);
    console.log('Data:', data);

    let response = await fetch(url)
    .then(data => {
        return data;
    })           

    const user = await response.json() 
}