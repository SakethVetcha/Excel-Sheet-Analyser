const { spawn } = require('child_process');
const path = require('path');

exports.handler = async function(event, context) {
  // Set up Python environment
  const python = spawn('python3', [
    path.join(__dirname, 'flask_app.py'),
    JSON.stringify(event)
  ]);

  return new Promise((resolve, reject) => {
    let result = '';
    let error = '';

    python.stdout.on('data', (data) => {
      result += data.toString();
    });

    python.stderr.on('data', (data) => {
      error += data.toString();
    });

    python.on('close', (code) => {
      if (code !== 0) {
        resolve({
          statusCode: 500,
          body: JSON.stringify({ error: error || 'Python process failed' })
        });
        return;
      }

      try {
        const response = JSON.parse(result);
        resolve({
          statusCode: response.statusCode || 200,
          headers: response.headers || {
            'Content-Type': 'application/json'
          },
          body: response.body || result
        });
      } catch (e) {
        resolve({
          statusCode: 200,
          headers: {
            'Content-Type': 'text/html'
          },
          body: result
        });
      }
    });

    python.on('error', (err) => {
      resolve({
        statusCode: 500,
        body: JSON.stringify({ error: err.message })
      });
    });
  });
}; 