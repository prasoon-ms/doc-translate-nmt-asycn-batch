import { AuthenticatedTemplate, UnauthenticatedTemplate, useMsal } from "@azure/msal-react";
import { API_SCOPE } from "./auth";
import { UploadForm } from "./components/UploadForm";
import { JobList } from "./components/JobList";

function SignInButton() {
  const { instance } = useMsal();
  return (
    <button
      onClick={async () => {
        const result = await instance.loginPopup({ scopes: [API_SCOPE] });
        instance.setActiveAccount(result.account);
      }}
    >
      Sign in with Microsoft
    </button>
  );
}

function SignOutButton() {
  const { instance, accounts } = useMsal();
  return (
    <button onClick={() => instance.logoutPopup()}>
      Sign out ({accounts[0]?.username})
    </button>
  );
}

export function App() {
  return (
    <main style={{ fontFamily: "system-ui, sans-serif", maxWidth: 720, margin: "2rem auto", padding: "0 1rem" }}>
      <h1>Document Translator</h1>
      <UnauthenticatedTemplate>
        <p>Please sign in to translate documents.</p>
        <SignInButton />
      </UnauthenticatedTemplate>
      <AuthenticatedTemplate>
        <SignOutButton />
        <hr />
        <UploadForm />
        <hr />
        <JobList />
      </AuthenticatedTemplate>
    </main>
  );
}
