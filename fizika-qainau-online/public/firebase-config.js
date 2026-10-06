/* Firebase жобасының веб-баптаулары.
   Firebase console → Project settings → General → Your apps → Web app → SDK setup and configuration → Config
   мәндерін төмендегі орындарға қойыңыз.

   Бұл мәндер құпия емес: олар тек жобаны атайды. Деректерді қорғайтын — firestore.rules
   ережелері және мұғалімнің Firebase Authentication-дағы құпия сөзі (ол бұл файлда да, HTML-де де жоқ). */
window.FIREBASE_CONFIG = {
  apiKey: 'YOUR_API_KEY',
  authDomain: 'YOUR_PROJECT_ID.firebaseapp.com',
  projectId: 'YOUR_PROJECT_ID',
  storageBucket: 'YOUR_PROJECT_ID.firebasestorage.app',
  messagingSenderId: 'YOUR_SENDER_ID',
  appId: 'YOUR_APP_ID'
};
