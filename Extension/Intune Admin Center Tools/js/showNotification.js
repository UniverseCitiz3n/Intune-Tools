function showNotification(message, type = 'success') {
  const notificationBar = document.getElementById('notification-bar');
  if (notificationBar) {
    notificationBar.textContent = message;
    notificationBar.className = type;
    notificationBar.style.display = 'block';
    setTimeout(() => {
      notificationBar.style.display = 'none';
    }, 3000);
  } else {
    console.warn('Notification bar not found, falling back to console:', message);
    console.log(message);
  }
}