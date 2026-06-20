export default async function handler(req, res) {
  if (req.method === 'POST') {
    try {
      const { monto, token, email } = req.body;
      
      res.status(200).json({
        success: true,
        message: 'Pago procesado correctamente',
        transactionId: 'TXN-' + Date.now(),
        monto: monto
      });
    } catch (error) {
      res.status(500).json({
        success: false,
        error: error.message
      });
    }
  } else {
    res.status(405).json({ error: 'Metodo no permitido' });
  }
}
