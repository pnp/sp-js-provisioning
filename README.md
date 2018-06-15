![Office 365 Developer Patterns and Practices](https://camo.githubusercontent.com/a732087ed949b0f2f84f5f02b8c79f1a9dd96f65/687474703a2f2f692e696d6775722e636f6d2f6c3031686876452e706e67)

# JavaScript Provisioning Library

This is the new home of code that started in the [core js library](https://github.com/SharePoint/PnP-JS-Core). It made sense to break it out so 
it can grow and develop independently. What is here has been heavily refactored to take advantage of the sp-pnp-js library so that it can work on nodejs.

Moving forward the [PnP core team](https://dev.office.com/patterns-and-practices) is looking for folks who are interested in developing the provisioning capabilities in nodejs to 
work with us to realize that goal. Our recommendation remains using the [CSOM based provisioning engine](https://github.com/SharePoint/PnP-Sites-Core) and calling it from 
client script as a microservice. Alternatively you can make use of the [PowerShell wrappers](https://github.com/SharePoint/PnP-powershell) to automate your site provisioning.

This library should currently be considered a work in progress

### Get Started

**NPM**

Add the npm packages to your project

    npm install sp-pnp-js sp-pnp-provisioning --save

### Authors
This project's contributors include Microsoft and [community contributors](AUTHORS). Work is done as as open source community project.

### Code of Conduct
This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

### "Sharing is Caring"

### Disclaimer
**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**
