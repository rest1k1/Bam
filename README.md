test("token: статус 200", function () {
  expect(res.getStatus()).to.equal(200);
});

const body = res.getBody();

test("token: поле token присутствует", function () {
  expect(body).to.have.property("token");
});

test("token: token — непустая строка", function () {
  expect(body.token).to.be.a("string").and.not.empty;
});

test("token: похож на JWT", function () {
  expect(body.token.split(".").length).to.equal(3);
});

// КЛЮЧЕВОЕ: сохраняем токен в переменную коллекции
bru.setVar("token", body.token);

// Запоминаем имя, чтобы сверить в /hello
bru.setVar("sentUsername", bru.getVar("username"));
