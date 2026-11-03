const input = document.querySelector('#document');
const label = document.querySelector('#file-label');
const button = document.querySelector('#convert-button');
const form = document.querySelector('#convert-form');

input.addEventListener('change', () => {
  if (input.files && input.files.length > 0) {
    label.textContent = input.files[0].name;
  }
});

form.addEventListener('submit', () => {
  button.disabled = true;
  button.textContent = 'Converting…';
});
