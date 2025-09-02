import { useState } from 'react';
     import axios from 'axios';
     import './App.css';

     function App() {
       const [deliveryDate, setDeliveryDate] = useState('>2025-09-02');
       const [data, setData] = useState([]);
       const [error, setError] = useState(null);

       const handleSubmit = async () => {
         try {
           setError(null);
           const response = await axios.post('http://localhost:8000/export-csv', {
             delivery_date: deliveryDate
           });
           setData(response.data.data);
         } catch (error) {
           setError(error.response?.data?.detail || 'Request failed');
           console.error('Error:', error);
         }
       };

       return (
         <div>
           <h1>JobRunner CSV Query</h1>
           <input
             type="text"
             value={deliveryDate}
             onChange={(e) => setDeliveryDate(e.target.value)}
             placeholder="Delivery Date (e.g., >2025-09-02)"
           />
           <button onClick={handleSubmit}>Query</button>
           {error && <p style={{ color: 'red' }}>{error}</p>}
           {data.length > 0 && (
             <pre>{JSON.stringify(data, null, 2)}</pre>
           )}
         </div>
       );
     }

     export default App;