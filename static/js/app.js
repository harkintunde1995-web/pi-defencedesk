/* DWF Ireland — Document Generator — Frontend JS */

document.addEventListener('DOMContentLoaded', function () {

  // ─── COLLAPSE TOGGLE ICON ───
  document.querySelectorAll('.section-header').forEach(function (header) {
    var targetId = header.getAttribute('data-bs-target');
    var target = document.querySelector(targetId);
    if (!target) return;

    target.addEventListener('show.bs.collapse', function () {
      header.querySelector('.toggle-icon').style.transform = 'rotate(0deg)';
    });
    target.addEventListener('hide.bs.collapse', function () {
      header.querySelector('.toggle-icon').style.transform = 'rotate(-90deg)';
    });
  });

  // ─── WRC FIELDS ───
  var courtSelect = document.querySelector('select[name="court_type"]');
  function toggleWrcFields() {
    var isWrc = courtSelect && courtSelect.value.toLowerCase().includes('wrc');
    document.querySelectorAll('.wrc-field').forEach(function (el) {
      el.style.display = isWrc ? '' : 'none';
    });
  }
  if (courtSelect) {
    courtSelect.addEventListener('change', toggleWrcFields);
    toggleWrcFields();
  }

  // ─── SELECT ALL / CLEAR ALL ───
  document.getElementById('selectAll').addEventListener('click', function () {
    document.querySelectorAll('input[name="documents"]').forEach(function (cb) {
      cb.checked = true;
    });
  });
  document.getElementById('clearAll').addEventListener('click', function () {
    document.querySelectorAll('input[name="documents"]').forEach(function (cb) {
      cb.checked = false;
    });
  });

  // ─── FORM SUBMIT ───
  var form = document.getElementById('mainForm');
  var generateBtn = document.getElementById('generateBtn');
  var btnText = document.getElementById('btnText');
  var btnSpinner = document.getElementById('btnSpinner');
  var errorAlert = document.getElementById('errorAlert');

  form.addEventListener('submit', function (e) {
    e.preventDefault();

    // Validate at least one document selected
    var selected = document.querySelectorAll('input[name="documents"]:checked');
    if (selected.length === 0) {
      showError('Please select at least one document to generate.');
      return;
    }

    // Validate required fields
    var missing = [];
    form.querySelectorAll('[required]').forEach(function (field) {
      if (!field.value.trim()) {
        missing.push(field.closest('.col-md-4, .col-md-6, .col-md-8, .col-12')
          ?.querySelector('.form-label')?.textContent?.replace('*', '').trim() || field.name);
      }
    });
    if (missing.length > 0) {
      showError('Please fill in the required fields: ' + missing.slice(0, 3).join(', ') +
        (missing.length > 3 ? ' and ' + (missing.length - 3) + ' more.' : '.'));
      return;
    }

    hideError();
    setLoading(true);

    var formData = new FormData(form);

    fetch('/generate', {
      method: 'POST',
      body: formData
    })
    .then(function (response) {
      if (!response.ok) {
        return response.json().then(function (data) {
          throw new Error(data.error || 'Generation failed.');
        });
      }
      return response.blob();
    })
    .then(function (blob) {
      // Trigger download
      var url = URL.createObjectURL(blob);
      var a = document.createElement('a');
      a.href = url;
      a.download = 'DWF_Documents_' + Date.now() + '.zip';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      setLoading(false);
    })
    .catch(function (err) {
      showError(err.message || 'An error occurred. Please try again.');
      setLoading(false);
    });
  });

  function setLoading(loading) {
    generateBtn.disabled = loading;
    btnText.style.display = loading ? 'none' : '';
    btnSpinner.style.display = loading ? '' : 'none';
  }

  function showError(msg) {
    errorAlert.textContent = msg;
    errorAlert.style.display = '';
    errorAlert.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }

  function hideError() {
    errorAlert.style.display = 'none';
  }

});
