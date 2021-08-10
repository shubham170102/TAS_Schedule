import firebase from 'firebase'

var firebaseConfig = {
    apiKey: "AIzaSyBzgZ18CDudRqJ_uMt9RGqh1F4onWszWjo",
    authDomain: "react-firebase-a4d09.firebaseapp.com",
    projectId: "react-firebase-a4d09",
    storageBucket: "react-firebase-a4d09.appspot.com",
    messagingSenderId: "545062225613",
    appId: "1:545062225613:web:27a02e678706c194bc4b84",
    measurementId: "G-DTWMG587NL"
  };

  var fire = firebase.initializeApp(firebaseConfig);

  export default fire;