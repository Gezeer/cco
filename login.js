const usuariosPermitidos = [
  {
    usuario: "admin",
    senha: "123456",
    perfil: "Administrador"
  },
  {
    usuario: "cco",
    senha: "1234",
    perfil: "Operador"
  },
  {
    usuario: "diretoria",
    senha: "2025",
    perfil: "Diretoria"
  }
];

document.addEventListener("DOMContentLoaded", () => {
  const usuarioSalvo = localStorage.getItem("usuarioLembrado");

  if (usuarioSalvo) {
    document.getElementById("usuario").value = usuarioSalvo;
    document.getElementById("lembrarUsuario").checked = true;
  }

  document.addEventListener("keydown", event => {
    if (event.key === "Enter") {
      entrar();
    }
  });
});

function entrar() {
  const usuario = document.getElementById("usuario").value.trim();
  const senha = document.getElementById("senha").value.trim();
  const erro = document.getElementById("erro");
  const lembrar = document.getElementById("lembrarUsuario").checked;

  erro.textContent = "";

  if (!usuario || !senha) {
    erro.textContent = "Preencha usuário e senha.";
    return;
  }

  const encontrado = usuariosPermitidos.find(
    item => item.usuario === usuario && item.senha === senha
  );

  if (!encontrado) {
    erro.textContent = "Usuário ou senha inválidos.";
    return;
  }

  localStorage.setItem("usuarioLogado", JSON.stringify(encontrado));

  if (lembrar) {
    localStorage.setItem("usuarioLembrado", usuario);
  } else {
    localStorage.removeItem("usuarioLembrado");
  }

  window.location.href = "index.html";
}