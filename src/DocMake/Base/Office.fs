namespace DocMake.Base

module Office = 

    let refobj (x:'a) : ref<obj> = ref (x :> obj)

