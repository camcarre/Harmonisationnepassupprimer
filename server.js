const https = require('https');
const fs = require('fs');
const path = require('path');
const xml2js = require('xml2js');
const selfsigned = require('selfsigned');

const PORT = 5890;

// Fonction pour mettre à jour le manifest avec le port fixe
async function updateManifest() {
    const manifestPath = path.join(__dirname, 'manifest.xml');
    const manifestContent = fs.readFileSync(manifestPath, 'utf-8');
    
    const parser = new xml2js.Parser();
    const builder = new xml2js.Builder();
    
    try {
        const result = await parser.parseStringPromise(manifestContent);
        
        // Mettre à jour les URLs dans le manifest
        function updateUrls(obj) {
            for (let key in obj) {
                if (typeof obj[key] === 'object') {
                    updateUrls(obj[key]);
                } else if (typeof obj[key] === 'string' && obj[key].includes('localhost')) {
                    obj[key] = obj[key].replace(/localhost:\d+/, `localhost:${PORT}`);
                }
            }
        }
        
        updateUrls(result);
        
        // Sauvegarder le manifest mis à jour
        const updatedXml = builder.buildObject(result);
        fs.writeFileSync(manifestPath, updatedXml);
        
        console.log(`Manifest mis à jour avec le port ${PORT}`);
    } catch (err) {
        console.error('Erreur lors de la mise à jour du manifest:', err);
    }
}

// Fonction pour créer un certificat auto-signé
function createCertificate() {
    const certPath = path.join(__dirname, 'server.crt');
    const keyPath = path.join(__dirname, 'server.key');
    
    if (!fs.existsSync(certPath) || !fs.existsSync(keyPath)) {
        console.log('Création du certificat auto-signé...');
        const attrs = [{ name: 'commonName', value: 'localhost' }];
        const pems = selfsigned.generate(attrs, {
            algorithm: 'sha256',
            days: 365,
            keySize: 2048
        });
        
        fs.writeFileSync(keyPath, pems.private);
        fs.writeFileSync(certPath, pems.cert);
        console.log('Certificat créé avec succès');
    }
    
    return {
        key: fs.readFileSync(keyPath),
        cert: fs.readFileSync(certPath)
    };
}

// Démarrage du serveur
async function startServer() {
    try {
        await updateManifest();
        
        const credentials = createCertificate();
        
        const server = https.createServer(credentials, (req, res) => {
            console.log(`Requête reçue: ${req.method} ${req.url}`);
            
            // Gérer les requêtes CORS
            res.setHeader('Access-Control-Allow-Origin', '*');
            res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
            res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
            
            if (req.method === 'OPTIONS') {
                res.writeHead(200);
                res.end();
                return;
            }
            
            let filePath = path.join(__dirname, req.url === '/' ? 'taskpane.html' : req.url.split('?')[0]);
            
            // Vérifier si le fichier existe
            if (fs.existsSync(filePath)) {
                const ext = path.extname(filePath);
                const contentType = {
                    '.html': 'text/html; charset=utf-8',
                    '.js': 'application/javascript; charset=utf-8',
                    '.css': 'text/css; charset=utf-8',
                    '.png': 'image/png',
                    '.jpg': 'image/jpeg',
                    '.svg': 'image/svg+xml',
                    '.xml': 'application/xml; charset=utf-8',
                    '.json': 'application/json; charset=utf-8',
                    '.map': 'application/json; charset=utf-8',
                    '.txt': 'text/plain; charset=utf-8'
                }[ext] || 'application/octet-stream';
                
                console.log(`Fichier trouvé: ${filePath}`);
                console.log(`Type MIME: ${contentType}`);
                
                res.writeHead(200, { 'Content-Type': contentType });
                fs.createReadStream(filePath).pipe(res);
            } else {
                console.log(`Fichier non trouvé: ${filePath}`);
                res.writeHead(404);
                res.end(`Fichier non trouvé: ${req.url}`);
            }
        });
        
        server.listen(PORT, () => {
            console.log(`\nServeur démarré sur https://localhost:${PORT}`);
            console.log('\nPour installer l\'add-in dans Excel :');
            console.log('1. Ouvrez Excel');
            console.log('2. Fichier > Options > Centre de gestion de la confidentialité');
            console.log('3. Paramètres du Centre de gestion de la confidentialité > Catalogues de compléments approuvés');
            console.log(`4. Ajoutez cette URL : https://localhost:${PORT}`);
            console.log('5. Cochez "Afficher dans le menu"');
            console.log('6. Redémarrez Excel\n');
        });
        
    } catch (err) {
        console.error('Erreur lors du démarrage du serveur:', err);
    }
}

startServer(); 