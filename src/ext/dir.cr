class Dir
  def self.mktmpdir : String
    path = File.join(Dir.tempdir, "#{Time.now.to_unix}-#{Random.rand(0x100000000).to_s(36)}")
    Dir.mkdir(path, 0o0700)
    path
  end
end
