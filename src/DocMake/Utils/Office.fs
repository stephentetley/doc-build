namespace DocMake.Utils

module Office = 

    let refobj (x:'a) : ref<obj> = ref (x :> obj)

