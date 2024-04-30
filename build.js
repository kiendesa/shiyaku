const builder = require('electron-builder');
const Platform = builder.Platform;

builder.build({
    targets: Platform.WINDOWS.createTarget(),
    config: {
        'productName': 'App_year',
        'copyright': 'Copyright 2024 Densan project.',
        'appId': 'com.example.my-electron-app',
        'win': {
            'target': {
                'target': 'zip',
                'arch': [
                    'x64',
                    'ia32',
                ]
            }
        }
    }
})
    .then(() => {
        console.log('\n***** Build-process is finished *****\n');
    })
    .catch((error) => {
        console.log(error.message);
    });
