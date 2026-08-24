import express from 'express';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';
import chatHandler from './api/chat.js';
import driveHandler from './api/drive.js';

const app = express();
const __dirname = dirname(fileURLToPath(import.meta.url));

app.use(express.json());
app.use(express.static(__dirname));

app.post('/api/chat', chatHandler);
app.post('/api/drive', driveHandler);

app.get('*', (req, res) => res.sendFile(join(__dirname, 'index.html')));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Skyscale running on port ${PORT}`));
