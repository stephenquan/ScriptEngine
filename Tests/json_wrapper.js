function parse(str)
{
  try {
    return JSON.parse(str);
  } catch (err) {
    return {
      f: "json_parse",
      args: str,
      error: err,
      message: JSON.stringify(err)
    }
  }
}

function stringify(obj)
{
  return JSON.stringify(obj, undefined, 2);
}


