# Office Add-in Command Example

This project demonstrates an example implementation of an Office Add-in command using the Office JavaScript API. The add-in showcases how to execute a specific action and display a notification message within an Office application.

## Getting Started

1. Clone the repository to your local development environment.
2. Install the required dependencies by running `npm install`.
3. Compile the TypeScript code using the command `npm run build`.
4. Sideload the add-in in an Office application, such as Outlook, to test its functionality.

## Project Structure

The project consists of the following key files:

- `commands.html`: The HTML file responsible for loading the Office JavaScript API.
- `commands.ts`: The TypeScript file containing the logic for the add-in command.
- `tsconfig.json`: The TypeScript configuration file specifying compiler options.

## Add-in Functionality

1. The `commands.html` file includes the Office JavaScript API from the CDN, ensuring the necessary dependencies are loaded.
2. The `commands.ts` file defines the `action` function, which is triggered when the add-in command is executed.
3. Inside the `action` function:
   - A notification message object is created, specifying the message type, content, icon, and persistence settings.
   - The `Office.context.mailbox.item.notificationMessages.replaceAsync` method is used to display the notification message in the Office application.
   - The `event.completed()` method is called to signal the completion of the add-in command function.
4. The `action` function is registered with Office using the `Office.actions.associate` method.

## TypeScript Configuration

The `tsconfig.json` file provides the necessary TypeScript compiler options for this project. It includes settings such as:
- Allowing JavaScript files (`allowJs`)
- Specifying the base URL for module resolution (`baseUrl`)
- Enabling compatibility with ES module import/export syntax (`esModuleInterop`)
- Enabling experimental decorators (`experimentalDecorators`)
- Setting the JSX syntax to React (`jsx`)
- Emitting errors if the output cannot be generated (`noEmitOnError`)
- Specifying the output directory for compiled files (`outDir`)
- Generating source map files (`sourceMap`)
- Targeting ECMAScript 5 (`target`)
- Including ES2015 and DOM libraries (`lib`)
- Excluding specific directories from compilation (`exclude`)
- Enabling file-based compilation in `ts-node` (`ts-node.files`)

For more detailed information on developing Office Add-ins and utilizing the Office JavaScript API, please refer to the official [Office Add-ins documentation](https://docs.microsoft.com/en-us/office/dev/add-ins/).