import './App.css';
import LoginButton from './components/LoginButton';
import LogoutButton from './components/LogoutButton';
import Profile from './components/Profile';
//import  {useAuth0}  from '@auth0/auth0-react';
import { Component } from 'react';
import fire from './fire';


class App extends Component {
  state = {
      text: ""
  }

  //isLoading = useAuth0();

  handleText=e=>{
        this.setState({text:e.target.value});
    }

    handleSubmit=e=>{
        //let messageRef = fire.database().ref('messages').orderByKey().limitToLast(100);
        fire.database().ref('messages').push(this.state.text);
        this.setState({
            text : ""
        });
    }

  render () { 
    //const isLoading = useAuth0();

        return (
        <>
        <LoginButton />
        <LogoutButton />
        <Profile />
        <br/>
                <input type='text'  onChange={this.handleText} id='inputText' value={this.state.text}/>
                <br/>
                <button onClick={this.handleSubmit}>Upload</button>
        </>
      );
  }
}

export default App;


//const MyDiv = () => {
 // console.log('loading...')
  //const value = useAuth0();
  ///return value;
//}