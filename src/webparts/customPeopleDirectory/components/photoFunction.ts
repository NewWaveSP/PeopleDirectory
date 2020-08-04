async function appendUserPhoto(client, userId) {
    try {
        let usrRes = await client.api('/users/{' + userId + '}/photo/$value').get();
        const blobUrl = window.URL.createObjectURL(new Blob(usrRes, { type: "image/jpeg" }));
        return blobUrl;
    } catch (e) {
        let blobUrl = "https://oromia.sharepoint.com/sites/NWT/SiteAssets/Images/avatar.PNG";
        return blobUrl;
    }
}

export { appendUserPhoto };