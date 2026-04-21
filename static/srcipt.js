// Selector de prioridad visual
document.querySelectorAll('.prioridad-option').forEach(label => {
    label.addEventListener('click', function() {
        const container = this.parentElement;
        container.querySelectorAll('.prioridad-option').forEach(l => l.classList.remove('seleccionada'));
        this.classList.add('seleccionada');
        this.querySelector('input[type="radio"]').checked = true;
    });
});

// Actualizar progreso (se puede mejorar con AJAX, pero por ahora es estático)
console.log('✅ Gestor de Tareas PRO cargado');