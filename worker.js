import { pipeline, env } from 'https://cdn.jsdelivr.net/npm/@huggingface/transformers@3.1.2/dist/transformers.min.js';

env.allowLocalModels = false;
let generator = null;

async function loadModel() {
    if (generator !== null) return generator;

    try {
        self.postMessage({ status: 'loading', message: 'Iniciando descarga de Qwen2.5-0.5B...' });
        
        generator = await pipeline(
            'text-generation',
            'onnx-community/Qwen2.5-0.5B-Instruct',
            {
                dtype: 'q4',
                device: 'webgpu',
                progress_callback: (data) => {
                    self.postMessage({
                        status: 'loading',
                        message: `${data.file || 'archivo'} - ${Math.round(data.progress || 0)}%`
                    });
                }
            }
        );
        
        self.postMessage({ status: 'ready' });
        return generator;
    } catch (error) {
        self.postMessage({ status: 'error', message: error.message });
        throw error;
    }
}

// Empezar a cargar en background
loadModel();

self.addEventListener('message', async (event) => {
    if (event.data.action !== 'analyze') return;
    
    try {
        const pipe = await loadModel();
        const stats = event.data.stats;
        
        // Construir el prompt con los datos
        const promptText = `
        Datos de ventas:
        - Ventas Totales: $${stats.totalSales}
        - Ventas con Promoción: $${stats.promoSales} (${stats.promoPct}%)
        - Incremento promedio (Lift): +${stats.lift}%
        - Promoción más exitosa: ${stats.bestPromo}
        `;

        const messages = [
            { 
                role: 'system', 
                content: 'Eres un analista de datos experto y hablas español. Basándote en las estadísticas proporcionadas, escribe 3 conclusiones breves (bullet points) sobre el impacto de las promociones. No agregues introducciones ni des explicaciones largas. Usa un tono corporativo.'
            },
            { role: 'user', content: promptText }
        ];
        
        self.postMessage({ status: 'generating' });
        
        const output = await pipe(messages, {
            max_new_tokens: 200,
            temperature: 0.2,
            do_sample: false
        });
        
        const generatedText = output[0].generated_text;
        
        const lastMessage = Array.isArray(generatedText) 
            ? generatedText[generatedText.length - 1].content 
            : generatedText;
            
        self.postMessage({ 
            status: 'complete', 
            output: lastMessage.trim() 
        });
        
    } catch (error) {
        self.postMessage({ status: 'error', message: error.message });
    }
});
