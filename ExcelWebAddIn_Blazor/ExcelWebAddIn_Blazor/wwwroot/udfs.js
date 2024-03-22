(function () {
    Office.initialize = function (reason) {
        console.log(reason, "office initialized");
    }
})();

async function sayHelloWorldAsync() {
    if (Excel) {
        Excel.run(function (ctx) {
            var rng = ctx.workbook.worksheets.getActiveWorksheet().getRange("A1")
            return ctx.sync()
                .then(function (val) {
                    rng.value = "Hello world"
                })
                .then(ctx.sync)
        });
    }
    else {
        return "Hello world"
    }
}

function sayHi() {
    return "Hi! Excel"
}