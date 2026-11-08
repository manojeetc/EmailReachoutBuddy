// Inline "change email" editor shared by the dashboard and the professor page.
document.addEventListener('click', e => {
  const btn = e.target.closest('[data-email-edit]');
  if (!btn) return;
  const row = btn.closest('.emailrow');
  const link = row.querySelector('.email, .emaillink');
  if (!link || row.querySelector('input')) return;
  const input = document.createElement('input');
  input.type = 'email';
  input.className = 'emailinput';
  input.value = link.textContent.trim();
  input.setAttribute('aria-label', 'Email address');
  link.hidden = true;
  btn.hidden = true;
  row.prepend(input);
  input.focus();
  input.select();
  let done = false;
  const close = () => { input.remove(); link.hidden = false; btn.hidden = false; };
  const finish = async save => {
    if (done) return;
    done = true;
    const value = input.value.trim();
    if (!save || !value || value === link.textContent.trim()) return close();
    const fd = new FormData();
    fd.append('csrf_token', document.querySelector('meta[name=csrf]').content);
    fd.append('email', value);
    try {
      const r = await fetch(btn.dataset.save, { method: 'POST', body: fd });
      const d = await r.json();
      if (!r.ok) throw new Error(d.error || 'Could not save the email.');
      link.textContent = d.email;
      link.href = 'mailto:' + d.email;
      link.title = d.email;
      close();
      if (document.getElementById('mail-form')) location.reload();
    } catch (err) { alert(err.message); close(); }
  };
  input.addEventListener('blur', () => finish(true));
  input.addEventListener('keydown', ev => {
    if (ev.key === 'Enter') { ev.preventDefault(); input.blur(); }
    else if (ev.key === 'Escape') finish(false);
  });
});
