const { ipcRenderer } = require('electron')

process.once('loaded', () => {
  window.addEventListener('message', evt => {
    if (evt.data.type === 'select-dirs') {
      ipcRenderer.send('select-dirs')
    }
  })
})

window.addEventListener('DOMContentLoaded', () => {
  ipcRenderer.on('change-path', (_event, value) => {
    const pathText = document.getElementById("path-text");
    pathText.value = value;
    const event = new Event("input");
    pathText.dispatchEvent(event);
  })
})
