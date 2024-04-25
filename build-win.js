const builder = require('electron-builder');

builder.build({
    config: {
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
});
