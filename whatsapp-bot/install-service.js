var Service = require('node-windows').Service;

var svc = new Service({
  name: 'WareCore WhatsApp Bot',
  description: 'WareCore Reports WhatsApp Bot (документы приёмки, без групп)',
  script: 'C:\\Users\\Administrator\\Desktop\\асер\\WareCore Reports\\whatsapp-bot\\bot.js',
  nodeOptions: ['--max-old-space-size=1024']
});

svc.on('install', function () {
  svc.start();
  console.log('Service installed and started');
});

svc.on('alreadyinstalled', function () {
  console.log('Service is already installed');
});

svc.on('start', function () {
  console.log('Service started');
});

svc.on('stop', function () {
  console.log('Service stopped');
});

svc.install();
