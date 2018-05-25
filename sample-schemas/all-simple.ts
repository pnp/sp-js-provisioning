import Schema from "../src/schema";

let template: Schema = {

    CustomActions: [{
        Name: "Test",
        Title: "Test Title",
        Location: "EditControlBlock",
        Url: "https://someurl"
    }],

    Features: [{
        id: "87294c72-f260-42f3-a41b-981a2ffce37a",
        deactivate: true,
        force: false
    }],

    WebSettings: {
        TreeViewEnabled: true
    },

    Navigation: {
        QuickLaunch: [
            {
                Title: "First One",
                Url: "https://bing.com"
            },
            {
                Title: "Second One",
                Url: "https://bing.com",
                Children: [{
                    Title: "Child - 1",
                    Url: "https://bing.com"
                },
                {
                    Title: "Child - 2",
                    Url: "https://bing.com"
                }]
            }
        ],
        TopNavigationBar: [
            {
                Title: "First One",
                Url: "https://bing.com"
            },
            {
                Title: "Second One",
                Url: "https://bing.com",
                Children: [{
                    Title: "Child!",
                    Url: "https://bing.com"
                }]
            }
        ]
    },

    Lists: [{
        Title: "Prov List 1",
        ContentTypesEnabled: true,
        Description: "This is the description",
        Template: 100
    },
    {
        Title: "Documents",
        ContentTypesEnabled: true,
        Description: "This is the description",
        Template: 101,
        AdditionalSettings: {
            Description: "I've updated the description",
            EnableFolderCreation: true,
            ForceCheckout: true
        }
    }]
};

export default template;