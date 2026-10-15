import { runOojsonSample } from '../../../examples/oojson-sample.mjs';
const status = document.querySelector('#status');
function convert() {
  try {
    const result = runOojsonSample(document.querySelector('#text').value);
    document.querySelector('#xml').textContent = result.xml;
    document.querySelector('#json').textContent = JSON.stringify(result.output, null, 2);
    status.textContent = 'Passed: converted JSON → XML → JSON and preserved paragraph text.';
  } catch (error) {
    status.textContent = `Conversion failed: ${error.message}`;
  }
}
document.querySelector('#convert').addEventListener('click', convert);
convert();
