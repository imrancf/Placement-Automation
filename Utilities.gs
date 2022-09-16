function includeFile_(output, name, params) {
  var template = HtmlService.createTemplateFromFile(name);
  if (params) {
    Object.keys(params).forEach(function (key) {
      template[key] = params[key];
    })
  }
  output.$out.append(template.evaluate().getContent());
}
