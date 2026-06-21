module.exports = (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  
  if (req.method === 'OPTIONS') {
    res.writeHead(200);
    res.end();
    return;
  }
  
  if (req.method !== 'POST') {
    res.status(405).json({ error: 'Metodo no permitido' });
    return;
  }
  
  const { amount, token } = req.body;
  
  if (!amount || !token) {
    res.status(400).json({ error: 'Faltan parametros: amount, token' });
    return;
  }
  
  res.status(200).json({
    success: true,
    message: 'Pago procesado correctamente',
    amount: amount,
    status: 'pending_confirmation'
  });
};
